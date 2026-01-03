"""
Internal: provide AES support for pypdf's fallback crypto provider.

pypdf supports encrypted PDFs, but when neither `cryptography` nor `pycryptodome`
is installed it falls back to a provider that raises `DependencyError` for AES.
Some PDFs are encrypted with an empty password ("not password protected" for
end-users) and should still be readable.

This module patches pypdf at runtime to enable AES-ECB/CBC and the `CryptAES`
implementation using a small pure-Python AES implementation.
"""

from __future__ import annotations

import secrets
from typing import Iterable


def _pkcs7_pad(data: bytes, block_size: int) -> bytes:
    padding = block_size - (len(data) % block_size)
    return data + bytes([padding]) * padding


def _pkcs7_unpad(data: bytes, block_size: int) -> bytes:
    if not data:
        return data
    padding = data[-1]
    if padding < 1 or padding > block_size:
        raise ValueError("Invalid PKCS#7 padding")
    if data[-padding:] != bytes([padding]) * padding:
        raise ValueError("Invalid PKCS#7 padding")
    return data[:-padding]


def _chunks(data: bytes, size: int) -> Iterable[bytes]:
    for i in range(0, len(data), size):
        yield data[i : i + size]


# AES S-box and inverse S-box (FIPS-197)
_SBOX: list[int] = [
    0x63,
    0x7C,
    0x77,
    0x7B,
    0xF2,
    0x6B,
    0x6F,
    0xC5,
    0x30,
    0x01,
    0x67,
    0x2B,
    0xFE,
    0xD7,
    0xAB,
    0x76,
    0xCA,
    0x82,
    0xC9,
    0x7D,
    0xFA,
    0x59,
    0x47,
    0xF0,
    0xAD,
    0xD4,
    0xA2,
    0xAF,
    0x9C,
    0xA4,
    0x72,
    0xC0,
    0xB7,
    0xFD,
    0x93,
    0x26,
    0x36,
    0x3F,
    0xF7,
    0xCC,
    0x34,
    0xA5,
    0xE5,
    0xF1,
    0x71,
    0xD8,
    0x31,
    0x15,
    0x04,
    0xC7,
    0x23,
    0xC3,
    0x18,
    0x96,
    0x05,
    0x9A,
    0x07,
    0x12,
    0x80,
    0xE2,
    0xEB,
    0x27,
    0xB2,
    0x75,
    0x09,
    0x83,
    0x2C,
    0x1A,
    0x1B,
    0x6E,
    0x5A,
    0xA0,
    0x52,
    0x3B,
    0xD6,
    0xB3,
    0x29,
    0xE3,
    0x2F,
    0x84,
    0x53,
    0xD1,
    0x00,
    0xED,
    0x20,
    0xFC,
    0xB1,
    0x5B,
    0x6A,
    0xCB,
    0xBE,
    0x39,
    0x4A,
    0x4C,
    0x58,
    0xCF,
    0xD0,
    0xEF,
    0xAA,
    0xFB,
    0x43,
    0x4D,
    0x33,
    0x85,
    0x45,
    0xF9,
    0x02,
    0x7F,
    0x50,
    0x3C,
    0x9F,
    0xA8,
    0x51,
    0xA3,
    0x40,
    0x8F,
    0x92,
    0x9D,
    0x38,
    0xF5,
    0xBC,
    0xB6,
    0xDA,
    0x21,
    0x10,
    0xFF,
    0xF3,
    0xD2,
    0xCD,
    0x0C,
    0x13,
    0xEC,
    0x5F,
    0x97,
    0x44,
    0x17,
    0xC4,
    0xA7,
    0x7E,
    0x3D,
    0x64,
    0x5D,
    0x19,
    0x73,
    0x60,
    0x81,
    0x4F,
    0xDC,
    0x22,
    0x2A,
    0x90,
    0x88,
    0x46,
    0xEE,
    0xB8,
    0x14,
    0xDE,
    0x5E,
    0x0B,
    0xDB,
    0xE0,
    0x32,
    0x3A,
    0x0A,
    0x49,
    0x06,
    0x24,
    0x5C,
    0xC2,
    0xD3,
    0xAC,
    0x62,
    0x91,
    0x95,
    0xE4,
    0x79,
    0xE7,
    0xC8,
    0x37,
    0x6D,
    0x8D,
    0xD5,
    0x4E,
    0xA9,
    0x6C,
    0x56,
    0xF4,
    0xEA,
    0x65,
    0x7A,
    0xAE,
    0x08,
    0xBA,
    0x78,
    0x25,
    0x2E,
    0x1C,
    0xA6,
    0xB4,
    0xC6,
    0xE8,
    0xDD,
    0x74,
    0x1F,
    0x4B,
    0xBD,
    0x8B,
    0x8A,
    0x70,
    0x3E,
    0xB5,
    0x66,
    0x48,
    0x03,
    0xF6,
    0x0E,
    0x61,
    0x35,
    0x57,
    0xB9,
    0x86,
    0xC1,
    0x1D,
    0x9E,
    0xE1,
    0xF8,
    0x98,
    0x11,
    0x69,
    0xD9,
    0x8E,
    0x94,
    0x9B,
    0x1E,
    0x87,
    0xE9,
    0xCE,
    0x55,
    0x28,
    0xDF,
    0x8C,
    0xA1,
    0x89,
    0x0D,
    0xBF,
    0xE6,
    0x42,
    0x68,
    0x41,
    0x99,
    0x2D,
    0x0F,
    0xB0,
    0x54,
    0xBB,
    0x16,
]

_INV_SBOX: list[int] = [
    0x52,
    0x09,
    0x6A,
    0xD5,
    0x30,
    0x36,
    0xA5,
    0x38,
    0xBF,
    0x40,
    0xA3,
    0x9E,
    0x81,
    0xF3,
    0xD7,
    0xFB,
    0x7C,
    0xE3,
    0x39,
    0x82,
    0x9B,
    0x2F,
    0xFF,
    0x87,
    0x34,
    0x8E,
    0x43,
    0x44,
    0xC4,
    0xDE,
    0xE9,
    0xCB,
    0x54,
    0x7B,
    0x94,
    0x32,
    0xA6,
    0xC2,
    0x23,
    0x3D,
    0xEE,
    0x4C,
    0x95,
    0x0B,
    0x42,
    0xFA,
    0xC3,
    0x4E,
    0x08,
    0x2E,
    0xA1,
    0x66,
    0x28,
    0xD9,
    0x24,
    0xB2,
    0x76,
    0x5B,
    0xA2,
    0x49,
    0x6D,
    0x8B,
    0xD1,
    0x25,
    0x72,
    0xF8,
    0xF6,
    0x64,
    0x86,
    0x68,
    0x98,
    0x16,
    0xD4,
    0xA4,
    0x5C,
    0xCC,
    0x5D,
    0x65,
    0xB6,
    0x92,
    0x6C,
    0x70,
    0x48,
    0x50,
    0xFD,
    0xED,
    0xB9,
    0xDA,
    0x5E,
    0x15,
    0x46,
    0x57,
    0xA7,
    0x8D,
    0x9D,
    0x84,
    0x90,
    0xD8,
    0xAB,
    0x00,
    0x8C,
    0xBC,
    0xD3,
    0x0A,
    0xF7,
    0xE4,
    0x58,
    0x05,
    0xB8,
    0xB3,
    0x45,
    0x06,
    0xD0,
    0x2C,
    0x1E,
    0x8F,
    0xCA,
    0x3F,
    0x0F,
    0x02,
    0xC1,
    0xAF,
    0xBD,
    0x03,
    0x01,
    0x13,
    0x8A,
    0x6B,
    0x3A,
    0x91,
    0x11,
    0x41,
    0x4F,
    0x67,
    0xDC,
    0xEA,
    0x97,
    0xF2,
    0xCF,
    0xCE,
    0xF0,
    0xB4,
    0xE6,
    0x73,
    0x96,
    0xAC,
    0x74,
    0x22,
    0xE7,
    0xAD,
    0x35,
    0x85,
    0xE2,
    0xF9,
    0x37,
    0xE8,
    0x1C,
    0x75,
    0xDF,
    0x6E,
    0x47,
    0xF1,
    0x1A,
    0x71,
    0x1D,
    0x29,
    0xC5,
    0x89,
    0x6F,
    0xB7,
    0x62,
    0x0E,
    0xAA,
    0x18,
    0xBE,
    0x1B,
    0xFC,
    0x56,
    0x3E,
    0x4B,
    0xC6,
    0xD2,
    0x79,
    0x20,
    0x9A,
    0xDB,
    0xC0,
    0xFE,
    0x78,
    0xCD,
    0x5A,
    0xF4,
    0x1F,
    0xDD,
    0xA8,
    0x33,
    0x88,
    0x07,
    0xC7,
    0x31,
    0xB1,
    0x12,
    0x10,
    0x59,
    0x27,
    0x80,
    0xEC,
    0x5F,
    0x60,
    0x51,
    0x7F,
    0xA9,
    0x19,
    0xB5,
    0x4A,
    0x0D,
    0x2D,
    0xE5,
    0x7A,
    0x9F,
    0x93,
    0xC9,
    0x9C,
    0xEF,
    0xA0,
    0xE0,
    0x3B,
    0x4D,
    0xAE,
    0x2A,
    0xF5,
    0xB0,
    0xC8,
    0xEB,
    0xBB,
    0x3C,
    0x83,
    0x53,
    0x99,
    0x61,
    0x17,
    0x2B,
    0x04,
    0x7E,
    0xBA,
    0x77,
    0xD6,
    0x26,
    0xE1,
    0x69,
    0x14,
    0x63,
    0x55,
    0x21,
    0x0C,
    0x7D,
]


def _xtime(a: int) -> int:
    a &= 0xFF
    return ((a << 1) ^ 0x1B) & 0xFF if (a & 0x80) else (a << 1) & 0xFF


def _gf_mul(a: int, b: int) -> int:
    result = 0
    a &= 0xFF
    b &= 0xFF
    while b:
        if b & 1:
            result ^= a
        a = _xtime(a)
        b >>= 1
    return result & 0xFF


def _add_round_key(state: list[int], round_key: bytes) -> None:
    for i in range(16):
        state[i] ^= round_key[i]


def _sub_bytes(state: list[int]) -> None:
    for i in range(16):
        state[i] = _SBOX[state[i]]


def _inv_sub_bytes(state: list[int]) -> None:
    for i in range(16):
        state[i] = _INV_SBOX[state[i]]


def _shift_rows(state: list[int]) -> None:
    for row in range(1, 4):
        row_bytes = [state[row + 4 * col] for col in range(4)]
        row_bytes = row_bytes[row:] + row_bytes[:row]
        for col in range(4):
            state[row + 4 * col] = row_bytes[col]


def _inv_shift_rows(state: list[int]) -> None:
    for row in range(1, 4):
        row_bytes = [state[row + 4 * col] for col in range(4)]
        row_bytes = row_bytes[-row:] + row_bytes[:-row]
        for col in range(4):
            state[row + 4 * col] = row_bytes[col]


def _mix_columns(state: list[int]) -> None:
    for col in range(4):
        i = 4 * col
        a0, a1, a2, a3 = state[i : i + 4]
        state[i + 0] = _gf_mul(a0, 2) ^ _gf_mul(a1, 3) ^ a2 ^ a3
        state[i + 1] = a0 ^ _gf_mul(a1, 2) ^ _gf_mul(a2, 3) ^ a3
        state[i + 2] = a0 ^ a1 ^ _gf_mul(a2, 2) ^ _gf_mul(a3, 3)
        state[i + 3] = _gf_mul(a0, 3) ^ a1 ^ a2 ^ _gf_mul(a3, 2)


def _inv_mix_columns(state: list[int]) -> None:
    for col in range(4):
        i = 4 * col
        a0, a1, a2, a3 = state[i : i + 4]
        state[i + 0] = (
            _gf_mul(a0, 14) ^ _gf_mul(a1, 11) ^ _gf_mul(a2, 13) ^ _gf_mul(a3, 9)
        )
        state[i + 1] = (
            _gf_mul(a0, 9) ^ _gf_mul(a1, 14) ^ _gf_mul(a2, 11) ^ _gf_mul(a3, 13)
        )
        state[i + 2] = (
            _gf_mul(a0, 13) ^ _gf_mul(a1, 9) ^ _gf_mul(a2, 14) ^ _gf_mul(a3, 11)
        )
        state[i + 3] = (
            _gf_mul(a0, 11) ^ _gf_mul(a1, 13) ^ _gf_mul(a2, 9) ^ _gf_mul(a3, 14)
        )


def _rcon(n: int) -> list[int]:
    rcon = [0] * (n + 1)
    rcon[1] = 0x01
    for i in range(2, n + 1):
        rcon[i] = _xtime(rcon[i - 1])
    return rcon


def _rot_word(word: list[int]) -> list[int]:
    return word[1:] + word[:1]


def _sub_word(word: list[int]) -> list[int]:
    return [_SBOX[b] for b in word]


def _expand_key(key: bytes) -> list[bytes]:
    if len(key) not in (16, 24, 32):
        raise ValueError("Invalid AES key length")

    nk = len(key) // 4
    nr = nk + 6
    w: list[list[int]] = [list(key[4 * i : 4 * i + 4]) for i in range(nk)]
    rcon = _rcon(nr)

    for i in range(nk, 4 * (nr + 1)):
        temp = w[i - 1][:]
        if i % nk == 0:
            temp = _sub_word(_rot_word(temp))
            temp[0] ^= rcon[i // nk]
        elif nk > 6 and i % nk == 4:
            temp = _sub_word(temp)
        w.append([a ^ b for a, b in zip(w[i - nk], temp)])

    round_keys: list[bytes] = []
    for r in range(nr + 1):
        key_bytes: list[int] = []
        for word in w[4 * r : 4 * r + 4]:
            key_bytes.extend(word)
        round_keys.append(bytes(key_bytes))
    return round_keys


def _aes_encrypt_block(block: bytes, round_keys: list[bytes]) -> bytes:
    if len(block) != 16:
        raise ValueError("Invalid AES block size")
    nr = len(round_keys) - 1
    state = list(block)

    _add_round_key(state, round_keys[0])
    for r in range(1, nr):
        _sub_bytes(state)
        _shift_rows(state)
        _mix_columns(state)
        _add_round_key(state, round_keys[r])
    _sub_bytes(state)
    _shift_rows(state)
    _add_round_key(state, round_keys[nr])
    return bytes(state)


def _aes_decrypt_block(block: bytes, round_keys: list[bytes]) -> bytes:
    if len(block) != 16:
        raise ValueError("Invalid AES block size")
    nr = len(round_keys) - 1
    state = list(block)

    _add_round_key(state, round_keys[nr])
    for r in range(nr - 1, 0, -1):
        _inv_shift_rows(state)
        _inv_sub_bytes(state)
        _add_round_key(state, round_keys[r])
        _inv_mix_columns(state)
    _inv_shift_rows(state)
    _inv_sub_bytes(state)
    _add_round_key(state, round_keys[0])
    return bytes(state)


def aes_ecb_encrypt(key: bytes, data: bytes) -> bytes:
    if len(data) % 16 != 0:
        raise ValueError("AES ECB requires data length multiple of 16")
    round_keys = _expand_key(key)
    return b"".join(
        _aes_encrypt_block(block, round_keys) for block in _chunks(data, 16)
    )


def aes_ecb_decrypt(key: bytes, data: bytes) -> bytes:
    if len(data) % 16 != 0:
        raise ValueError("AES ECB requires data length multiple of 16")
    round_keys = _expand_key(key)
    return b"".join(
        _aes_decrypt_block(block, round_keys) for block in _chunks(data, 16)
    )


def aes_cbc_encrypt(key: bytes, iv: bytes, data: bytes) -> bytes:
    if len(iv) != 16:
        raise ValueError("AES CBC requires 16-byte IV")
    if len(data) % 16 != 0:
        raise ValueError("AES CBC requires data length multiple of 16")
    round_keys = _expand_key(key)
    out = bytearray()
    prev = iv
    for block in _chunks(data, 16):
        xored = bytes(b ^ p for b, p in zip(block, prev))
        enc = _aes_encrypt_block(xored, round_keys)
        out.extend(enc)
        prev = enc
    return bytes(out)


def aes_cbc_decrypt(key: bytes, iv: bytes, data: bytes) -> bytes:
    if len(iv) != 16:
        raise ValueError("AES CBC requires 16-byte IV")
    if len(data) % 16 != 0:
        raise ValueError("AES CBC requires data length multiple of 16")
    round_keys = _expand_key(key)
    out = bytearray()
    prev = iv
    for block in _chunks(data, 16):
        dec = _aes_decrypt_block(block, round_keys)
        out.extend(bytes(b ^ p for b, p in zip(dec, prev)))
        prev = block
    return bytes(out)


def patch_pypdf_fallback_aes() -> bool:
    """
    Patch pypdf to support AES when running on the fallback crypto provider.

    This overwrite is slow in absolute runtime behavior but is still pure-python without any
    large c-dependencies.

    Returns:
        True if patching was applied, False otherwise.
    """
    import pypdf._crypt_providers as providers

    if providers.crypt_provider[0] != "local_crypt_fallback":
        return False

    import pypdf._crypt_providers._fallback as fb
    import pypdf._encryption as enc

    # Patch the fallback provider module.
    fb.aes_ecb_encrypt = aes_ecb_encrypt
    fb.aes_ecb_decrypt = aes_ecb_decrypt
    fb.aes_cbc_encrypt = aes_cbc_encrypt
    fb.aes_cbc_decrypt = aes_cbc_decrypt

    def _cryptaes_init(self: object, key: bytes) -> None:
        setattr(self, "key", key)

    def _cryptaes_encrypt(self: object, data: bytes) -> bytes:
        iv = secrets.token_bytes(16)
        padded = _pkcs7_pad(data, 16)
        return iv + aes_cbc_encrypt(getattr(self, "key"), iv, padded)

    def _cryptaes_decrypt(self: object, data: bytes) -> bytes:
        iv = data[:16]
        payload = data[16:]
        if not payload:
            return payload
        if len(payload) % 16 != 0:
            payload = _pkcs7_pad(payload, 16)
        plain = aes_cbc_decrypt(getattr(self, "key"), iv, payload)
        return _pkcs7_unpad(plain, 16)

    fb.CryptAES.__init__ = _cryptaes_init  # type: ignore[method-assign]
    fb.CryptAES.encrypt = _cryptaes_encrypt  # type: ignore[method-assign]
    fb.CryptAES.decrypt = _cryptaes_decrypt  # type: ignore[method-assign]

    # Patch the exported bindings that pypdf modules imported earlier.
    providers.aes_ecb_encrypt = aes_ecb_encrypt
    providers.aes_ecb_decrypt = aes_ecb_decrypt
    providers.aes_cbc_encrypt = aes_cbc_encrypt
    providers.aes_cbc_decrypt = aes_cbc_decrypt
    providers.CryptAES = fb.CryptAES  # type: ignore[assignment]

    enc.aes_ecb_encrypt = aes_ecb_encrypt
    enc.aes_ecb_decrypt = aes_ecb_decrypt
    enc.aes_cbc_encrypt = aes_cbc_encrypt
    enc.aes_cbc_decrypt = aes_cbc_decrypt
    enc.CryptAES = fb.CryptAES  # type: ignore[assignment]

    return True
