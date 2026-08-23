# Iterator Semantics

This document defines the public iteration contract for `read_file`,
`read_bytes`, and `read_many`.

## Public contract

All three functions return a synchronous, lazy, single-pass
`Iterator[ExtractedDocument]`. The contract deliberately promises `Iterator`,
not the concrete generator implementation. Callers should therefore depend on
normal iteration behavior rather than generator-only methods.

An iterator is used even for formats that normally produce one document
because containers such as archives and `.mbox` files can produce an arbitrary
number of documents.

## Validation and deferred work

Configuration that can be checked without processing source content is
validated when the public function is called. I/O and parsing remain deferred
until the returned iterator is advanced.

| API | Validated when called | Deferred until iteration |
|---|---|---|
| `read_file` | Path, source size, routing, and ZIP-bomb configuration | Opening, reading, parsing, and normalization |
| `read_bytes` | Input type and size, routing hints, and ZIP-bomb configuration | Stream positioning, parsing, and normalization |
| `read_many` | Folder, selection, result-callback, and ZIP-bomb configuration | Traversal and each selected file's validation and extraction |

For `read_many`, a selected file's own validation is deferred because folder
contents are discovered lazily.

## Consumption

Each returned iterator can be consumed only once. Once exhausted, it remains
exhausted. Call the extraction function again to repeat extraction.

Advancing an iterator can perform file I/O, decompression, parsing, and
normalization. A yielded `ExtractedDocument` is fully materialized, so its
fields remain usable after the iterator advances.

Use normal iteration when a source may yield any number of documents:

```python
for document in sharepoint2text.read_file("mailbox.mbox"):
    consume(document)
```

When exactly one document is an application invariant, iterable unpacking
checks that invariant and exhausts the iterator:

```python
document, = sharepoint2text.read_file("report.pdf")
```

This raises `ValueError` for zero or multiple documents. `next(...)` returns
only the first document and does not verify exhaustion.

## Ordering

Documents from one source are yielded in the order supplied by its extractor.
For `read_many`, all documents from one selected top-level file are adjacent,
and the next file is not processed until the current file is exhausted.

Filesystem traversal order is not a portable sorting guarantee. Applications
that require a stable global order must impose one themselves, accepting the
enumeration or buffering cost that entails.

## Errors and partial results

`read_file` and `read_bytes` propagate lazy extraction failures to the caller.

`read_many` catches recoverable `ExtractionError` and I/O failures for an
individual selected file, logs the failure, and continues with later files.
Documents yielded before such a failure remain valid and are not retracted.
Unexpected exceptions stop iteration.

### Structured per-file reporting

`read_many` accepts an optional `on_file_result` callback. It receives one
`BatchFileResult` after a selected top-level file either:

- completes successfully, or
- encounters a recoverable extraction or I/O failure.

The result contains:

- `path`: the selected top-level file path;
- `documents_extracted`: documents yielded from that file before completion or
  failure;
- `error`: the recoverable exception, or `None` on success; and
- `succeeded`: a convenience property equivalent to `error is None`.

The outer iterator must be advanced past a file's final document before the
success callback runs. No callback is delivered for filtered files, for the
current file when iteration is abandoned early, or for an unexpected exception.
A callback exception propagates and stops batch iteration.

The library invokes callbacks as files finish and does not retain their
results. This preserves bounded library-side reporting memory even when a
folder contains arbitrarily many files. A caller that appends every result to
a list chooses memory proportional to the number of processed files.

## Resource lifetime

Full exhaustion releases open file, archive, and parser resources. Breaking
early can retain resources until the concrete iterator is closed or finalized.

The current implementation uses generators with `close()` at runtime, but the
stable public API promises only `Iterator`. Deterministic context-managed
cleanup after partial consumption is not currently part of the contract.

## Memory behavior

Laziness bounds the library's result retention: extraction processes one
top-level source at a time and does not accumulate all yielded documents or
per-file reports. It does not guarantee that parsing a single source uses
constant memory; individual extractors may materialize source content.

Calling `list(...)` explicitly retains every yielded document and removes the
iterator's result-memory benefit.

## Non-goals

The current iterator API does not promise:

- replayability or random access;
- asynchronous iteration;
- concurrent file extraction;
- portable filesystem ordering;
- constant-memory parsing within a single source;
- automatic retention of a complete batch report; or
- deterministic cleanup through a public context-manager protocol.
