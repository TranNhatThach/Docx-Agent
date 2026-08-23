"""
Native Word Track Changes Integration:
Maps Semantic ChangeObjects to Microsoft Word OpenXML revision markup (<w:ins>, <w:del>, <w:pPrChange>).
"""

from typing import List, Dict, Any, Optional
from datetime import datetime
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from docx_agent.changes.model import ChangeObject, ChangeType, ChangeStatus


class NativeWordRevisionExporter:
    """
    Applies OpenXML Track Changes markup into a Word document package from semantic changes.
    """

    @classmethod
    def apply_word_revisions(
        cls,
        doc: docx.Document,
        changes: List[ChangeObject],
        author: str = "Antigravity AI Agent",
    ) -> None:
        """
        Injects standard OOXML revision tags so Microsoft Word shows native Track Changes balloons.
        """
        now_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

        for idx, chg in enumerate(changes):
            if chg.status != ChangeStatus.PROPOSED and chg.status != ChangeStatus.ACCEPTED:
                continue

            rev_id = str(idx + 100)

            # Find matching paragraph in docx if possible
            for p in doc.paragraphs:
                if chg.target_element in p.text or (chg.before_content and str(chg.before_content) in p.text):
                    if chg.change_type == ChangeType.REPLACE_TEXT and chg.before_content and chg.after_content:
                        # Wrap deletion and insertion
                        p_elem = p._p
                        ins_xml = parse_xml(
                            f'<w:ins {nsdecls("w")} w:id="{rev_id}" w:author="{author}" w:date="{now_iso}">'
                            f'<w:r><w:t>{chg.after_content}</w:t></w:r>'
                            f'</w:ins>'
                        )
                        del_xml = parse_xml(
                            f'<w:del {nsdecls("w")} w:id="{int(rev_id)+1}" w:author="{author}" w:date="{now_iso}">'
                            f'<w:r><w:delText>{chg.before_content}</w:delText></w:r>'
                            f'</w:del>'
                        )
                        p_elem.append(del_xml)
                        p_elem.append(ins_xml)
                    elif chg.change_type == ChangeType.INSERT_TEXT and chg.after_content:
                        p_elem = p._p
                        ins_xml = parse_xml(
                            f'<w:ins {nsdecls("w")} w:id="{rev_id}" w:author="{author}" w:date="{now_iso}">'
                            f'<w:r><w:t>{chg.after_content}</w:t></w:r>'
                            f'</w:ins>'
                        )
                        p_elem.append(ins_xml)
                    elif chg.change_type == ChangeType.DELETE_TEXT and chg.before_content:
                        p_elem = p._p
                        del_xml = parse_xml(
                            f'<w:del {nsdecls("w")} w:id="{rev_id}" w:author="{author}" w:date="{now_iso}">'
                            f'<w:r><w:delText>{chg.before_content}</w:delText></w:r>'
                            f'</w:del>'
                        )
                        p_elem.append(del_xml)
                    break

    @classmethod
    def export_docx_with_revisions(
        cls,
        doc_node: Any,
        changes: List[ChangeObject],
        output_path: Any,
        author: str = "Antigravity AI Agent",
    ) -> str:
        """
        Exports a DocumentNode into a .docx with native Track Changes markup.
        """
        from docx_agent.adapters.docx import DocxExporter
        path = DocxExporter.export_docx(doc_node, output_path)
        doc = docx.Document(str(path))
        cls.apply_word_revisions(doc, changes, author=author)
        doc.save(str(path))
        return str(path)

