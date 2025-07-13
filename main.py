from fastapi import FastAPI, Response, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, EmailStr
from typing import List, Optional
from docxtpl import DocxTemplate
from docx import Document
import io
import json
import math
import random
import httpx

def extract_audit_table_iso9001_stage1(docx_path_or_stream):
    doc = Document(docx_path_or_stream)
    extracted = {}

    for table in doc.tables:
        first_row = [cell.text.strip() for cell in table.rows[0].cells]
        if (
            "Clause & Description" in first_row
            and "Document Verification detail with statement of Conformity" in first_row
        ):
            second_row = [cell.text.strip() for cell in table.rows[1].cells]
            try:
                cl_no_idx = second_row.index("Cl. NO.")
                desc_idx = second_row.index("Description")
                doc_ver_idx = next(
                    i for i, h in enumerate(first_row)
                    if "Document Verification detail with statement of Conformity" in h
                )
            except ValueError:
                continue

            for row in table.rows[2:]:
                cells = row.cells
                if len(cells) <= max(cl_no_idx, desc_idx, doc_ver_idx):
                    continue
                cl_no = cells[cl_no_idx].text.strip()
                desc = cells[desc_idx].text.strip()
                doc_ver = cells[doc_ver_idx].text.strip()
                key = f"{cl_no} - {desc}" if cl_no and desc else cl_no or desc
                if key and doc_ver:
                    extracted[key] = (
                        extracted[key] + "\n" + doc_ver if key in extracted else doc_ver
                    )
            break
    return extracted

def replace_clause_placeholders_random(audit_dict, placeholder="{{clause}}"):
    keys_with_placeholder = [k for k, v in audit_dict.items() if v == placeholder]
    total = len(keys_with_placeholder)
    if total == 0:
        return audit_dict

    nc_count = min(2, max(1, math.floor(0.1 * total)))
    o_count = max(1, math.ceil(0.1 * total))
    c_count = total - nc_count - o_count

    replacements = (["C"] * c_count) + (["O"] * o_count) + (["NC"] * nc_count)
    random.shuffle(replacements)

    for i, k in enumerate(keys_with_placeholder):
        audit_dict[k] = replacements[i]
    return audit_dict

def replace_clause_placeholders_in_docx(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    placeholders = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if "{{clause}}" in run.text:
                            placeholders.append(run)

    total = len(placeholders)
    if total == 0:
        return docx_bytes

    nc_count = min(2, max(1, math.floor(0.1 * total)))
    o_count = max(1, math.ceil(0.1 * total))
    c_count = total - nc_count - o_count

    replacements = (["C"] * c_count) + (["O"] * o_count) + (["NC"] * nc_count)
    random.shuffle(replacements)

    for run, replacement in zip(placeholders, replacements):
        run.text = run.text.replace("{{clause}}", replacement)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------- MODELS ----------

class ISO9001Stage1Audit(BaseModel):
    organizationName: str
    address: str
    siteAddress: Optional[str] = ""
    numberOfEmployees: int
    emailId: EmailStr
    contactPerson: str
    telephoneFax: str
    scope: str
    riskCategory: str
    iafCode: str
    auditTeam: List[str]
    dateOfAudit: str
    auditObjective: str
    briefDescription: str
    quotedManDaysAdequate: str
    changeInEmployeeDetail: str
    changeInScope: str
    additionalInformation: str
    internalAuditFrequency: str
    dateOfLastInternalAudit: str
    managementReviewFrequency: str
    dateOfLastManagementReview: str
    recommendationForStage2: str
    reviewedBy: str
    dateOfReview: str
    clientName: str
    designation: str

# ---------- ROUTE ----------

@app.post("/iso9001/stage1/submit")
async def submit_iso9001_stage1(audit: ISO9001Stage1Audit):
    # Step 1: Render DOCX with original data
    doc = DocxTemplate("templates/iso9001_stage1.docx")
    context = {
        "organizationName": audit.organizationName,
        "address": audit.address,
        "siteAddress": audit.siteAddress,
        "numberOfEmployees": audit.numberOfEmployees,
        "emailId": audit.emailId,
        "contactPerson": audit.contactPerson,
        "telephoneFax": audit.telephoneFax,
        "scope": audit.scope,
        "riskCategory": audit.riskCategory,
        "iafCode": audit.iafCode,
        "auditTeam": "\n".join(audit.auditTeam),
        "dateOfAudit": audit.dateOfAudit,
        "auditObjective": audit.auditObjective,
        "briefDescription": audit.briefDescription,
        "quotedManDaysAdequate": audit.quotedManDaysAdequate,
        "changeInEmployeeDetail": audit.changeInEmployeeDetail,
        "changeInScope": audit.changeInScope,
        "additionalInformation": audit.additionalInformation,
        "internalAuditFrequency": audit.internalAuditFrequency,
        "dateOfLastInternalAudit": audit.dateOfLastInternalAudit,
        "managementReviewFrequency": audit.managementReviewFrequency,
        "dateOfLastManagementReview": audit.dateOfLastManagementReview,
        "recommendationForStage2": audit.recommendationForStage2,
        "reviewedBy": audit.reviewedBy,
        "dateOfReview": audit.dateOfReview,
        "clientName": audit.clientName,
        "designation": audit.designation,
        "clause": "{{clause}}",
    }
    doc.render(context)

    extract_buffer = io.BytesIO()
    doc.save(extract_buffer)
    extract_buffer.seek(0)

    extracted_data = extract_audit_table_iso9001_stage1(extract_buffer)
    print("Extracted Data:", json.dumps(extracted_data, indent=2, ensure_ascii=False))

    extracted_data = replace_clause_placeholders_random(extracted_data)

    prompt = f"""
    Rephrase the following audit content professionally while maintaining all technical details.
Keep the organization name exactly as provided below.
Rephrase and adapt the content so that it is contextually relevant and specific to the provided scope.
If any part of the template is not applicable to the scope, adjust or omit details accordingly, but keep the section key present.
Do not modify or remove any keys. Only rephrase and adapt the values.

Organization Name: {audit.organizationName}
Scope: {audit.scope}

Audit Content to rephrase and adapt:
{json.dumps(extracted_data, indent=2, ensure_ascii=False)}

Provide the rephrased and scope-adapted version in the same JSON structure with improved language.
Output ONLY the JSON object, without any additional explanation or text.

    """

    mistral_api_url = "https://mistral-api-5icm.onrender.com/api/mistral"
    payload = {
        "prompt": prompt,
    }
    headers = {
        "Content-Type": "application/json"
    }

    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(mistral_api_url, json=payload, headers=headers)
            response.raise_for_status()

            # Handle both JSON and streaming responses
            if response.headers.get("content-type") == "application/json":
                mistral_result = response.json()
                rephrased_text = mistral_result.get("generated_text", "")
            else:
                rephrased_text = response.text

        print("Mistral Response:", rephrased_text)

        try:
            rephrased_data = json.loads(rephrased_text)
            if isinstance(rephrased_data, dict):
                for key, value in rephrased_data.items():
                    if key in context:
                        context[key] = value
        except json.JSONDecodeError:
            context["briefDescription"] = rephrased_text.strip()

    except Exception as e:
        print(f"Error calling Mistral API: {str(e)}")
        rephrased_text = json.dumps(extracted_data, indent=2)

    doc = DocxTemplate("templates/iso9001_stage1.docx")
    doc.render(context)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)

    final_doc_bytes = replace_clause_placeholders_in_docx(output_buffer.getvalue())

    headers = {
        "Content-Disposition": f"attachment; filename={audit.organizationName}_iso9001_stage1_report.docx"
    }

    return Response(
        content=final_doc_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
