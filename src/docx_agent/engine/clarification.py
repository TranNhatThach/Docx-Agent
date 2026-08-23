"""
Clarification & Ambiguity Engine: Detects underspecified user instructions
and generates structured multiple-choice clarification requests when confidence is LOW.
"""

from enum import Enum
from typing import List, Dict, Any, Optional, Tuple
from pydantic import BaseModel, Field


class AgentConfidence(str, Enum):
    HIGH = "HIGH"        # Clear intent or safe defaults exist -> execute directly
    MEDIUM = "MEDIUM"    # Minor variation -> execute with explicit stated default
    LOW = "LOW"          # Material ambiguity -> prompt user for structured clarification


class ClarificationOption(BaseModel):
    id: str
    label: str
    description: Optional[str] = None
    inferred_parameters: Dict[str, Any] = Field(default_factory=dict)


class ClarificationRequest(BaseModel):
    question: str
    confidence: AgentConfidence = AgentConfidence.LOW
    options: List[ClarificationOption] = Field(default_factory=list)
    default_option_id: Optional[str] = None


class ClarificationEngine:
    """
    Evaluates instruction specificity against document context.
    """

    @staticmethod
    def assess_instruction(
        instruction: str,
        selected_text: Optional[str] = None,
        document_profile: str = "academic_report",
    ) -> Tuple[AgentConfidence, Optional[ClarificationRequest]]:
        """
        Assesses if an instruction has material ambiguity.
        """
        ins_lower = instruction.strip().lower()

        # Case 1: Generic rewrite request ("viết lại đoạn này", "rewrite this", "improve this")
        rewrite_triggers = ["viết lại", "rewrite", "chỉnh sửa đoạn", "làm hay hơn", "cải thiện đoạn", "tối ưu đoạn"]
        if any(phrase in ins_lower for phrase in rewrite_triggers):
            req = ClarificationRequest(
                question="Bạn muốn viết lại đoạn văn này theo phong cách nào?",
                confidence=AgentConfidence.LOW,
                options=[
                    ClarificationOption(
                        id="academic",
                        label="Học thuật & Trịnh trọng",
                        description="Sử dụng từ ngữ học thuật, cấu trúc câu chặt chẽ, khách quan.",
                        inferred_parameters={"tone": "academic", "expand": False},
                    ),
                    ClarificationOption(
                        id="concise",
                        label="Ngắn gọn & Súc tích",
                        description="Loại bỏ từ thừa, tập trung vào ý chính cốt lõi.",
                        inferred_parameters={"tone": "concise", "reduce_length": True},
                    ),
                    ClarificationOption(
                        id="technical",
                        label="Kỹ thuật chuyên sâu",
                        description="Nhấn mạnh thông số, thuật ngữ và cơ chế kỹ thuật.",
                        inferred_parameters={"tone": "technical", "detail": "high"},
                    ),
                ],
                default_option_id="academic" if document_profile == "academic_report" else "concise",
            )
            return AgentConfidence.LOW, req

        # Case 2: Generic image/diagram request ("thêm hình", "thêm diagram", "vẽ sơ đồ", "add image")
        diagram_triggers = ["thêm hình", "thêm diagram", "vẽ hình", "vẽ sơ đồ", "tạo sơ đồ", "tạo lưu đồ", "add image", "create diagram"]
        if any(phrase in ins_lower for phrase in diagram_triggers):
            req = ClarificationRequest(
                question="Bạn muốn chèn loại hình minh họa nào cho phần này?",
                confidence=AgentConfidence.LOW,
                options=[
                    ClarificationOption(
                        id="architecture",
                        label="Sơ đồ Kiến trúc Hệ thống (Architecture Diagram)",
                        description="Thể hiện các tầng (UI, Service, Database, Agent Engine).",
                        inferred_parameters={"diagram_type": "architecture"},
                    ),
                    ClarificationOption(
                        id="flowchart",
                        label="Lưu đồ Quy trình (Flowchart)",
                        description="Thể hiện luồng xử lý dữ liệu qua các bước.",
                        inferred_parameters={"diagram_type": "flowchart"},
                    ),
                    ClarificationOption(
                        id="use_case",
                        label="Sơ đồ Ca sử dụng (Use Case Diagram)",
                        description="Thể hiện tương tác giữa người dùng và hệ thống.",
                        inferred_parameters={"diagram_type": "use_case"},
                    ),
                ],
                default_option_id="architecture",
            )
            return AgentConfidence.LOW, req

        # Case 3: Instruction is specific -> High confidence
        return AgentConfidence.HIGH, None
