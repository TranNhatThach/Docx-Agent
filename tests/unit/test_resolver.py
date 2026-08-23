"""
Unit tests for TargetResolver and element identity.
"""

import pytest
from docx_agent.agent import DocumentAgent
from docx_agent.core.exceptions import ElementNotFoundError, AmbiguousTargetError


def test_resolve_by_id(sample_docx):
    agent = DocumentAgent(sample_docx)
    p = agent.resolver.resolve_paragraphs("p_0002", single=True)
    assert len(p) == 1
    assert "Giới thiệu" in p[0].text


def test_resolve_by_index(sample_docx):
    agent = DocumentAgent(sample_docx)
    p = agent.resolver.resolve_paragraphs("idx:1", single=True)
    assert len(p) == 1
    assert "Giới thiệu" in p[0].text


def test_resolve_by_text(sample_docx):
    agent = DocumentAgent(sample_docx)
    p = agent.resolver.resolve_paragraphs("đoạn văn mở đầu", single=True)
    assert len(p) == 1
    assert "giới thiệu về đề tài" in p[0].text

    multi = agent.resolver.resolve_paragraphs("Phương pháp", single=False)
    assert len(multi) == 2


def test_resolve_by_heading(sample_docx):
    agent = DocumentAgent(sample_docx)
    p = agent.resolver.resolve_paragraphs({"type": "heading", "level": 1}, single=True)
    assert "Giới thiệu" in p[0].text


def test_ambiguous_target_error(sample_docx):
    agent = DocumentAgent(sample_docx)
    # Both Heading 2 and body paragraph contain 'Phương pháp' or 'nghiên cứu'
    agent.append("Một nghiên cứu khác.")
    with pytest.raises(AmbiguousTargetError):
        agent.resolver.resolve_paragraphs("nghiên cứu", single=True)
