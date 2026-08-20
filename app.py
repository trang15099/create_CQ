import streamlit as st
from docxtpl import DocxTemplate
from datetime import date
from io import BytesIO
import re


st.set_page_config(
    page_title="CQ Generator",
    page_icon="📄",
    layout="wide"
)

st.title("CQ Generator - System")


# =========================
# FUNCTIONS
# =========================

def parse_serials(raw_text):
    """
    Cho phép Sales paste serial:
    - mỗi serial 1 dòng
    - cách nhau bằng dấu phẩy
    - cách nhau bằng space/tab
    """
    if not raw_text:
        return []

    serials = re.split(r"[\n,\t; ]+", raw_text.strip())
    return [x.strip() for x in serials if x.strip()]


def serial_to_rows(serials, columns=4):
    """
    Chia serial thành bảng 4 cột cho phụ lục
    """
    rows = []

    for i in range(0, len(serials), columns):
        row = serials[i:i + columns]

        while len(row) < columns:
            row.append("")

        rows.append(row)

    return rows


# =========================
# GENERAL INFORMATION
# =========================

st.subheader("1. Thông tin CQ")

col1, col2 = st.columns(2)

with col1:
    city = st.selectbox(
        "Thành phố cấp CQ",
        ["Hà Nội", "TP.HCM"]
    )

with col2:
    issue_date = st.date_input(
        "Ngày cấp CQ",
        value=date.today()
    )

eu_name = st.text_input(
    "Tên EU",
    placeholder="BAN QUẢN LÝ DỰ ÁN..."
)

address = st.text_input(
    "Địa chỉ",
    placeholder="Nhập địa chỉ EU"
)


# =========================
# PRODUCT INFORMATION
# =========================

st.divider()

st.subheader("2. Thông tin sản phẩm")

product_count = st.number_input(
    "Số loại sản phẩm",
    min_value=1,
    max_value=20,
    value=1,
    step=1
)

products = []

for i in range(product_count):

    with st.expander(
        f"Sản phẩm {i + 1}",
        expanded=True
    ):

        product_name = st.text_area(
            "Tên sản phẩm",
            key=f"product_name_{i}",
            height=100,
            placeholder="Ví dụ: MÁY TÍNH ĐỂ BÀN (PC) ASUS..."
        )

        col1, col2 = st.columns(2)

        with col1:
            quantity = st.number_input(
                "Số lượng",
                min_value=1,
                value=1,
                step=1,
                key=f"quantity_{i}"
            )

        with col2:
            origin = st.text_input(
                "Xuất xứ",
                value="Trung Quốc",
                key=f"origin_{i}"
            )

        serial_text = st.text_area(
            "Serial Number",
            key=f"serial_{i}",
            height=150,
            placeholder=(
                "Paste serial vào đây.\n"
                "Có thể mỗi serial 1 dòng hoặc copy trực tiếp từ Excel."
            )
        )

        serials = parse_serials(serial_text)

        if serials:
            st.caption(
                f"Đã nhận: {len(serials)} Serial Number"
            )

            if len(serials) <= 5:
                st.info(
                    "Serial sẽ được hiển thị trực tiếp trên CQ."
                )
            else:
                st.info(
                    'CQ chính sẽ ghi "Phụ lục đính kèm".'
                )

        products.append({
            "product_name": product_name,
            "quantity": quantity,
            "origin": origin,
            "serials": serials
        })


# =========================
# GENERATE CQ
# =========================

st.divider()

if st.button(
    "GENERATE CQ",
    type="primary",
    use_container_width=True
):

    # -------------------------
    # VALIDATION
    # -------------------------

    errors = []

    if not eu_name.strip():
        errors.append("Chưa nhập Tên EU.")

    if not address.strip():
        errors.append("Chưa nhập Địa chỉ.")

    for index, product in enumerate(products):

        if not product["product_name"].strip():
            errors.append(
                f"Sản phẩm {index + 1}: chưa nhập tên sản phẩm."
            )

        # Nếu muốn bắt buộc số serial = quantity,
        # bỏ comment đoạn dưới:

        # if len(product["serials"]) != product["quantity"]:
        #     errors.append(
        #         f"Sản phẩm {index + 1}: "
        #         f"Số lượng ({product['quantity']}) "
        #         f"không khớp số Serial ({len(product['serials'])})."
        #     )

    if errors:

        for error in errors:
            st.error(error)

        st.stop()


    # -------------------------
    # PROCESS DATA
    # -------------------------

    document_products = []
    appendix_products = []

    for product in products:

        serials = product["serials"]

        # <= 5: show directly
        if len(serials) <= 5:
            serial_display = "\n".join(serials)

        # >= 6: appendix
        else:
            serial_display = "Phụ lục đính kèm"

            appendix_products.append({
                "product_name": product["product_name"],
                "serial_rows": serial_to_rows(
                    serials,
                    columns=4
                )
            })

        document_products.append({
            "product_name": product["product_name"],
            "quantity": product["quantity"],
            "origin": product["origin"],
            "serial_display": serial_display
        })


    # -------------------------
    # CITY TEXT
    # -------------------------

    city_text = city

    # Nếu muốn giống chính xác wording template
    if city == "TP.HCM":
        city_text = "Tp.HCM"


    # -------------------------
    # TEMPLATE DATA
    # -------------------------

    context = {

        "city": city_text,

        "day": f"{issue_date.day:02d}",

        "month": f"{issue_date.month:02d}",

        "year": issue_date.year,

        "eu_name": eu_name.upper(),

        "address": address,

        "products": document_products,

        "has_appendix": len(appendix_products) > 0,

        "appendix_products": appendix_products
    }


    # -------------------------
    # CREATE WORD
    # -------------------------

    try:

        doc = DocxTemplate(
            "templates/CQ_SYSTEM_TEMPLATE.docx"
        )

        doc.render(context)

        output = BytesIO()

        doc.save(output)

        output.seek(0)

        # Clean filename
        safe_eu = re.sub(
            r'[\\/*?:"<>|]',
            "",
            eu_name
        )

        safe_eu = safe_eu[:40]

        filename = (
            f"CQ_SYSTEM_{safe_eu}_"
            f"{issue_date.strftime('%Y%m%d')}.docx"
        )

        st.success("CQ đã được tạo thành công.")

        st.download_button(
            label="DOWNLOAD WORD",
            data=output,
            file_name=filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True
        )

    except FileNotFoundError:

        st.error(
            "Không tìm thấy file CQ_SYSTEM_TEMPLATE.docx."
        )

    except Exception as e:

        st.error(
            f"Có lỗi khi tạo CQ: {str(e)}"
        )
