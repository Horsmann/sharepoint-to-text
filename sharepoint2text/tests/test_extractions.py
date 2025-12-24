import io
import logging
import unittest

from sharepoint2text.extractors.doc_extractor import read_doc
from sharepoint2text.extractors.docx_extractor import read_docx

logger = logging.getLogger(__name__)


def test_read_docx() -> None:
    filename = "sharepoint2text/tests/resources/GKIM_Skills_Framework_-_static.docx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    result = read_docx(file_like)
    test_case_obj = unittest.TestCase()

    expected = sorted(
        [
            "metadata",
            "paragraphs",
            "tables",
            "headers",
            "footers",
            "images",
            "hyperlinks",
            "footnotes",
            "endnotes",
            "comments",
            "sections",
            "styles",
            "full_text",
        ]
    )
    test_case_obj.assertListEqual(expected, sorted(result.keys()))

    # text is long. Verify only beginning and ending of the combined text
    test_case_obj.assertEqual(
        "Welcome to the Government", result["full_text"][:25].strip()
    )
    test_case_obj.assertEqual("improved records systems.", result["full_text"][-25:])

    test_case_obj.assertEqual(230, len(result["paragraphs"]))

    test_case_obj.assertListEqual(
        sorted(
            [
                "title",
                "author",
                "subject",
                "keywords",
                "category",
                "comments",
                "created",
                "modified",
                "last_modified_by",
                "revision",
            ]
        ),
        sorted(result["metadata"].keys()),
    )


def test_read_doc() -> None:
    with open(
        "sharepoint2text/tests/resources/Speech_Prime_Minister_of_The_Netherlands_EN.doc",
        mode="rb",
    ) as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    content = read_doc(file_like=file_like)

    test_case_obj = unittest.TestCase()

    # Text content
    expected = """
    Welcome by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende, at the Inaugural Session of the International Criminal Court, The Hague, 11 March 2003 \n\n(Check against delivery)\n\nYour Royal Highnesses, Secretary-General, Your Excellencies, ladies and gentlemen,\n\nA very warm welcome to The Hague, the heart of Dutch democracy. The Netherlands is proud to be your host. \n\nAnd a special welcome to today’s eighteen most important people, who will shortly be sworn in as the first judges at the International Criminal Court. My sincere congratulations on your election.\n\nFour hundred and twenty years ago, the great legal thinker Hugo Grotius was born in Delft, less than ten kilometres from this spot. He was active in Dutch and European politics. \n\nFate did not smile on him. He fell victim to internal political conflicts, and was imprisoned in Loevestein castle. But he escaped by hiding in a chest of books. Dutch schoolchildren still love that story.\n\nGrotius fled to France, where he wrote the book that was to make him famous and which was translated into many languages: On the Law of War and Peace. \n\nIn it Grotius sets out his ideal: a system of international law, with clear agreements and procedures for countries to comply with. He believed that a system of this kind was necessary for international justice and stability.\n\nToday, ladies and gentlemen, nearly four centuries later, we move a step closer to that ideal. The International Criminal Court adds a crucial new element to the international legal system. \n\nIt makes it possible to prosecute the most serious crimes (genocide, crimes against humanity and war crimes) if they are not prosecuted at national level.\n\nSo today, the eleventh of March two thousand and three, is a historic day. Today the international community shows that it is still committed to justice, despite the many bloody conflicts and treaty violations we have seen since the Second World War.\n\nSuspicion and pessimism often dominate international politics. But today we are showing the world that there are also grounds for joy, optimism and hope.\n\nOf course, there is still a long way to go. We know that some countries are reluctant to sign up. The International Criminal Court is like a young swan. It needs time to grow bigger and stronger, then it can spread its wings and everyone will see it fly. Our work is not yet done. But with all of our help the ICC will succeed.\n\nMany people have been looking forward to this day. Many people have worked hard to bring it about. In particular, President Arthur Robinson of Trinidad and Tobago, who put the ICC onto the United Nations’ agenda in the late nineteen-eighties. And the UN Secretary-General, Kofi Annan, who did so much to speed up its establishment.\n\nI would also mention the judges and staff of other international courts, especially the Yugoslavia and Rwanda tribunals. Their experience has been and will be most valuable to the ICC.\n\nAnd finally I would mention the non-governmental organisations that have given their backing. Without your enthusiasm and support, it would all have taken far longer.\n\nThe Netherlands, and The Hague in particular, is honoured to be the ICC’s host. Since the first international peace conference was held here, over a century ago, The Hague has developed into the judicial capital of the world. We are proud of that.\n\nBut today, all of us can be proud.\n\nHugo Grotius’s last words were: “I have attempted much but achieved nothing”. \n\nToday we can say we have achieved something Grotius could only dream of: an international criminal court as part of an international legal order. And that takes us a big step closer to international justice.\n\nIt now gives me great pleasure to give the floor to the President of the Assembly of States Parties, His Royal Highness Prince Zeid Ra’ad Zeid al-Hussein.\n\nThank you.
    """
    test_case_obj.assertEqual(expected.strip(), content["text"])
    test_case_obj.assertListEqual(
        sorted(["text", "footnotes", "metadata", "headers_footers", "comments"]),
        sorted(content),
    )

    # Metadata
    test_case_obj.assertEqual(
        "Short dinner speech by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende",
        content["metadata"]["title"],
    )
    test_case_obj.assertEqual("Toby Screech", content["metadata"]["author"])
    test_case_obj.assertEqual("", content["metadata"]["keywords"])
    test_case_obj.assertEqual(580, content["metadata"]["num_words"])
    test_case_obj.assertEqual("2003-03-13T09:03:00", content["metadata"]["create_time"])
    test_case_obj.assertEqual(
        "2003-03-13T09:03:00", content["metadata"]["last_saved_time"]
    )
