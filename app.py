import streamlit as st
from docxtpl import DocxTemplate
from datetime import date
from io import BytesIO
import re


# =========================
# PAGE CONFIG
# =========================

st.set_page_config(
    page_title="CQ Generator",
    page_icon="📄",
    layout="wide"
)

st.title("Tạo CQ - ASUS")


# =========================
# FUNCTIONS
# =========================

def parse_serials(raw_text):
    """
    Cho phép Sales paste serial:
    - mỗi serial 1 dòng
    - cách nhau bằng dấu phẩy
    - cách nhau bằng dấu ;
    - cách nhau bằng space/tab
    """
    if not raw_text:
        return []

    serials = re.split(
        r"[\n,\t; ]+",
        raw_text.strip()
    )

    return [
        x.strip()
        for x in serials
        if x.strip()
    ]


def serial_to_rows(serials, columns=4):
    """
    Chia serial thành bảng 4 cột cho phụ lục.
    """
    rows = []

    for i in range(0, len(serials), columns):
        row = serials[i:i + columns]

        while len(row) < columns:
            row.append("")

        rows.append(row)

    return rows


# =========================
# SESSION STATE
# =========================

if "product_table" not in st.session_state:
    st.session_state.product_table = [
        {
            "Tên Sản Phẩm": "",
            "Số lượng": 1,
            "Xuất xứ (Double click nếu chọn Đài Loan)": "Trung Quốc"
        }
    ]


# =========================
# 1. GENERAL INFORMATION
# =========================

st.subheader("1. Thông tin CQ")

col1, col2 = st.columns(2)

with col1:
    city = st.selectbox(
        "Thành phố cấp CQ",
        ["TP.HCM","Hà Nội"]
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
# 2. PRODUCT INFORMATION
# =========================

st.divider()

st.subheader("2. Thông tin sản phẩm")

st.caption(
    "Thêm/xóa dòng trực tiếp trong bảng bên dưới."
)


product_table = st.data_editor(
    st.session_state.product_table,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Model name": st.column_config.TextColumn(
            "Model name",
            width="large",
            required=True,
            help="Nhập tên đầy đủ của sản phẩm"
        ),

        "Số lượng": st.column_config.NumberColumn(
            "Số lượng",
            min_value=1,
            step=1,
            width="small",
            required=True
        ),

        "Xuất xứ": st.column_config.SelectboxColumn(
            "Xuất xứ",
            options=[
                "Trung Quốc",
                "Đài Loan"
            ],
            default="Trung Quốc",
            width="medium",
            required=True
        ),
    },
    key="product_editor"
)


# =========================
# CLEAN PRODUCT ROWS
# =========================

clean_rows = []

for row in product_table:

    product_name = str(
        row.get("Model name", "")
    ).strip()

    if not product_name:
        continue

    quantity = row.get(
        "Số lượng",
        1
    )

    if quantity is None:
        quantity = 1

    quantity = int(quantity)

    origin = str(
        row.get(
            "Xuất xứ",
            "Trung Quốc"
        )
    ).strip()

    if not origin:
        origin = "Trung Quốc"

    clean_rows.append({
        "product_name": product_name,
        "quantity": quantity,
        "origin": origin
    })


# =========================
# SERIAL NUMBER SECTION
# =========================

products = []

if clean_rows:

    st.markdown("### Serial Number")

    st.caption(
        "Paste Serial Number riêng cho từng model. Mỗi dòng 1 Serial"
        "Có thể copy trực tiếp từ Excel."
    )


    for i, row in enumerate(clean_rows):

        st.markdown(
            f"**{i + 1}. {row['product_name']}**"
        )

        serial_text = st.text_area(
            "Serial Number",
            key=f"serial_{i}",
            height=120,
            placeholder=(
                "Ví dụ:\n"
                "W3PFAC009325112\n"
                "W4PFAC009720156\n"
                "W4PFAC009736158"
            ),
            label_visibility="collapsed"
        )

        serials = parse_serials(
            serial_text
        )

        if serials:

            st.caption(
                f"Đã nhận: {len(serials)} Serial Number"
            )


        products.append({
            "product_name": row["product_name"],
            "quantity": row["quantity"],
            "origin": row["origin"],
            "serials": serials
        })


        st.write("")


else:

    st.info(
        "Hãy nhập ít nhất 1 Model name "
        "trong bảng để nhập Serial Number."
    )


# =========================
# APPENDIX STATUS
# =========================

use_appendix_for_all = any(
    len(product["serials"]) >= 6
    for product in products
)


if use_appendix_for_all:

    st.info(
        "Có ít nhất 1 model từ 6 Serial Number trở lên. "
        "Toàn bộ Serial Number của CQ "
        "sẽ được đưa xuống phụ lục."
    )


# =========================
# SUMMARY
# =========================

if products:

    total_models = len(products)

    total_serials = sum(
        len(product["serials"])
        for product in products
    )

    st.caption(
        f"{total_models} model(s) | "
        f"{total_serials} Serial Number"
    )


# =========================
# GENERATE CQ
# =========================

st.divider()


if st.button(
    "GENERATE CQ",
    type="primary",
    use_container_width=True
):

    # =========================
    # VALIDATION
    # =========================

    errors = []


    if not eu_name.strip():
        errors.append(
            "Chưa nhập Tên EU."
        )


    if not address.strip():
        errors.append(
            "Chưa nhập Địa chỉ."
        )


    if not products:
        errors.append(
            "Chưa nhập thông tin sản phẩm."
        )


    for index, product in enumerate(products):

        if not product["product_name"].strip():

            errors.append(
                f"Sản phẩm {index + 1}: "
                "chưa nhập Model name."
            )


        if not product["origin"].strip():

            errors.append(
                f"Sản phẩm {index + 1}: "
                "chưa chọn Xuất xứ."
            )


        # ==========================================
        # OPTIONAL:
        # Bật đoạn dưới nếu muốn bắt buộc:
        # Số lượng = số Serial
        # ==========================================

        # if len(product["serials"]) != product["quantity"]:
        #
        #     errors.append(
        #         f"Sản phẩm {index + 1}: "
        #         f"Số lượng ({product['quantity']}) "
        #         f"không khớp số Serial "
        #         f"({len(product['serials'])})."
        #     )


    if errors:

        for error in errors:
            st.error(error)

        st.stop()


    # =========================
    # PROCESS DATA
    # =========================

    document_products = []

    appendix_products = []


    use_appendix_for_all = any(
        len(product["serials"]) >= 6
        for product in products
    )


    for product in products:

        serials = product["serials"]


        # ==========================================
        # Có 1 model >= 6 serial
        # => tất cả model xuống phụ lục
        # ==========================================

        if use_appendix_for_all:

            serial_display = (
                "Phụ lục đính kèm"
            )


            appendix_products.append({
                "product_name":
                    product["product_name"],

                "serial_rows":
                    serial_to_rows(
                        serials,
                        columns=4
                    )
            })


        # ==========================================
        # Không có model nào >= 6 serial
        # => hiển thị trực tiếp trên CQ
        # ==========================================

        else:

            serial_display = "\n".join(
                serials
            )


        document_products.append({
            "product_name":
                product["product_name"],

            "quantity":
                product["quantity"],

            "origin":
                product["origin"],

            "serial_display":
                serial_display
        })


    # =========================
    # CITY TEXT
    # =========================

    city_text = city


    if city == "TP.HCM":
        city_text = "Tp.HCM"


    # =========================
    # TEMPLATE DATA
    # =========================

    context = {

        "city":
            city_text,

        "day":
            f"{issue_date.day:02d}",

        "month":
            f"{issue_date.month:02d}",

        "year":
            issue_date.year,

        "eu_name":
            eu_name.upper(),

        "address":
            address,

        "products":
            document_products,

        "has_appendix":
            use_appendix_for_all,

        "appendix_products":
            appendix_products
    }


    # =========================
    # CREATE WORD
    # =========================

    try:

        doc = DocxTemplate(
            "CQ_SYSTEM_TEMPLATE_OPTIONAL_APPENDIX.docx"
        )


        doc.render(
            context
        )


        output = BytesIO()


        doc.save(
            output
        )


        output.seek(0)


        # =========================
        # FILE NAME
        # CQ_EU_YYYYMMDD.docx
        # =========================

        safe_eu = re.sub(
            r'[\\/*?:"<>|]',
            "",
            eu_name
        )

        safe_eu = safe_eu.strip()


        filename = (
            f"CQ_{safe_eu}_"
            f"{issue_date.strftime('%Y%m%d')}.docx"
        )


        # =========================
        # DOWNLOAD
        # =========================

        st.success(
            "CQ đã được tạo thành công."
        )


        st.download_button(
            label="DOWNLOAD WORD",
            data=output,
            file_name=filename,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True
        )


    # =========================
    # TEMPLATE NOT FOUND
    # =========================

    except FileNotFoundError:

        st.error(
            "Không tìm thấy file "
            "CQ_SYSTEM_TEMPLATE_OPTIONAL_APPENDIX.docx "
            "trên GitHub."
        )


    # =========================
    # OTHER ERROR
    # =========================

    except Exception as e:

        st.error(
            f"Có lỗi khi tạo CQ: {str(e)}"
        )
