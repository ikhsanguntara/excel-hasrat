from typing import Optional, List, Any
from pydantic import BaseModel, Field

class CheckpointItem(BaseModel):
    checkpoint_id: Optional[Any] = None
    area_id: Optional[Any] = None
    section_id: Optional[Any] = None
    sectiondtl_id: Optional[Any] = None
    checkpoint_name: Optional[str] = ""
    field_code: Optional[str] = ""
    seq_num: Optional[int] = 0
    is_mandatory: Optional[str] = "N"
    mstopt_id: Optional[Any] = None
    created_user: Optional[str] = None
    created_date: Optional[str] = None
    updated_user: Optional[str] = None
    updated_date: Optional[str] = None
    product_group: Optional[str] = ""
    area_name: Optional[str] = "Lain-lain"
    area_seq: Optional[Any] = None
    section_name: Optional[str] = "Umum"
    section_seq: Optional[Any] = None
    sectiondtl_name: Optional[str] = ""
    sectiondtl_seq: Optional[Any] = None
    check_user: Optional[str] = "-"
    check_num: Optional[Any] = None
    doc_num: Optional[str] = "-"
    start_doc_date: Optional[str] = ""
    doc_date: Optional[str] = ""
    end_doc_date: Optional[str] = ""
    horresult_id: Optional[Any] = None
    result: Optional[str] = ""
    type_document: Optional[Any] = None
    img_path: Optional[str] = None
    hor_id: Optional[Any] = None
    hordetail_id: Optional[Any] = None
    section_date: Optional[str] = ""
    section_user: Optional[str] = ""
    is_sectiondone: Optional[str] = ""
    is_areadone: Optional[str] = ""

    class Config:
        extra = "allow"
