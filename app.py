import streamlit as st
from docxtpl import DocxTemplate
from datetime import date
from io import BytesIO
import pandas as pd
import re
import os

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


# =========================
# PAGE CONFIG
# =========================

st.set_page_config(
    page_title="CQ Generator",
    page_icon="📄",
    layout="wide"
)

st.title("Tạo CQ")


# =========================
# GOOGLE SERVICES
# =========================

def get_drive_service():
    """
    Google Drive dùng OAuth của tài khoản Google thật.
    File upload sẽ sử dụng quota My Drive của tài khoản này.
    """

    credentials = Credentials(
        token=None,
        refresh_token=st.secrets["google_oauth"]["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=st.secrets["google_oauth"]["client_id"],
        client_secret=st.secrets["google_oauth"]["client_secret"],
        scopes=[
            "https://www.googleapis.com/auth/drive"
        ]
    )

    return build(
        "drive",
        "v3",
        credentials=credentials
    )


def get_sheets_service():
    """
    Google Sheet vẫn dùng Service Account.
    """

    credentials = (
        service_account
        .Credentials
        .from_service_account_info(
            dict(
                st.secrets[
                    "gcp_service_account"
                ]
            ),
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets"
            ]
        )
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials
    )


# =========================
# GOOGLE DRIVE FUNCTIONS
# =========================

def find_or_create_folder(
    drive_service,
    parent_id,
    folder_name
):

    safe_name = folder_name.replace(
        "'",
        "\\'"
    )

    query = (
        f"name = '{safe_name}' "
        f"and '{parent_id}' in parents "
        f"and mimeType = "
        f"'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )

    result = (
        drive_service
        .files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id,name)"
        )
        .execute()
    )

    folders = result.get(
        "files",
        []
    )

    if folders:
        return folders[0]["id"]

    folder_metadata = {
        "name": folder_name,
        "mimeType":
            "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }

    folder = (
        drive_service
        .files()
        .create(
            body=folder_metadata,
            fields="id"
        )
        .execute()
    )

    return folder["id"]


def upload_cq_to_drive(
    drive_service,
    word_bytes,
    filename,
    issue_date
):

    # Đây PHẢI là ID folder CQ GENERATED
    root_folder_id = (
        st.secrets["google"]["drive_folder_id"]
    )

    # CQ GENERATED / YYYY
    year_folder_id = find_or_create_folder(
        drive_service,
        root_folder_id,
        str(issue_date.year)
    )

    # CQ GENERATED / YYYY / Tháng MM
    month_folder_id = find_or_create_folder(
        drive_service,
        year_folder_id,
        f"Tháng {issue_date.month:02d}"
    )

    media = MediaIoBaseUpload(
        BytesIO(word_bytes),
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        resumable=False
    )

    file_metadata = {
        "name": filename,
        "parents": [month_folder_id]
    }

    uploaded_file = (
        drive_service
        .files()
        .create(
            body=file_metadata,
            media_body=media,
            fields="id,name,webViewLink"
        )
        .execute()
    )

    return uploaded_file["webViewLink"]


# =========================
# GOOGLE SHEET FUNCTIONS
# =========================

def append_tracking_row(
    sheets_service,
    creator_name,
    drive_link,
    issue_date,
    form_type,
    city,
    eu_name,
    address,
    products
):

    sheet_id = (
        st.secrets["google"]["sheet_id"]
    )

    sheet_name = (
        st.secrets["google"]["sheet_name"]
    )

    # =========================
    # NEXT STT
    # =========================

    result = (
        sheets_service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A2:A"
        )
        .execute()
    )

    existing_rows = result.get(
        "values",
        []
    )

    stt = len(existing_rows) + 1

    # =========================
    # PRODUCT SUMMARY
    # =========================

    product_summary = "\n".join(
        [
            (
                f"{i + 1}. "
                f"{product['product_name']} "
                f"| SL: {product['quantity']} "
                f"| {product['origin']}"
            )
            for i, product
            in enumerate(products)
        ]
    )

    # =========================
    # TRACKING ROW A:L
    # =========================

    row = [
        stt,
        # A - STT

        creator_name.strip(),
        # B - KD yêu cầu

        False,
        # C - Đẩy SA check

        "",
        # D - Request time

        "",
        # E - SA phản hồi

        (
            f'=HYPERLINK('
            f'"{drive_link}",'
            f'"Mở CQ")'
        ),
        # F - Link CQ

        issue_date.strftime(
            "%d/%m/%Y"
        ),
        # G - Ngày CQ

        form_type,
        # H - Loại CQ

        city,
        # I - Khu vực

        eu_name.upper().strip(),
        # J - Tên EU

        address.strip(),
        # K - Địa chỉ

        product_summary
        # L - Sản phẩm
    ]

    (
        sheets_service
        .spreadsheets()
        .values()
        .append(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A:L",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={
                "values": [row]
            }
        )
        .execute()
    )


# =========================
# GENERAL FUNCTIONS
# =========================

def parse_serials(raw_text):

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


def serial_to_rows(
    serials,
    columns=4
):

    rows = []

    for i in range(
        0,
        len(serials),
        columns
    ):

        row = serials[
            i:i + columns
        ]

        while len(row) < columns:
            row.append("")

        rows.append(row)

    return rows


# =========================
# KD YÊU CẦU
# =========================

creator_name = st.text_input(
    "KD yêu cầu",
    placeholder="Nhập tên KD yêu cầu CQ"
)


# =========================
# 1. CHỌN LOẠI FORM
# =========================

st.subheader(
    "1. Loại CQ"
)

form_type = st.radio(
    "Chọn Form",
    [
        "ASUS",
        "DGW"
    ],
    horizontal=True
)


# =========================
# 2. THÔNG TIN CQ
# =========================

st.divider()

st.subheader(
    "2. Thông tin CQ"
)

col1, col2 = st.columns(2)

with col1:

    city = st.selectbox(
        "Thành phố cấp CQ",
        [
            "TP.HCM",
            "Hà Nội"
        ]
    )

with col2:

    issue_date = st.date_input(
        "Ngày cấp CQ",
        value=date.today()
    )


eu_name = st.text_input(
    "Tên EU",
    placeholder=(
        "BAN QUẢN LÝ DỰ ÁN..."
    )
)

address = st.text_input(
    "Địa chỉ",
    placeholder=(
        "Nhập địa chỉ EU"
    )
)


# =========================
# TEMPLATE STATUS
# =========================

if form_type == "ASUS":

    st.caption(
        "Form đang chọn: ASUS"
    )

elif city == "Hà Nội":

    st.caption(
        "Form đang chọn: DGW - Hà Nội"
    )

else:

    st.caption(
        "Form đang chọn: DGW - TP.HCM"
    )


# =========================
# 3. THÔNG TIN SẢN PHẨM
# =========================

st.divider()

st.subheader(
    "3. Thông tin sản phẩm"
)

st.markdown(
    """
    <small>
    Thêm/xóa dòng trực tiếp trong bảng bên dưới.
    Nếu copy paste lưu ý xóa lỗi xuống dòng trong từng ô.<br>
    CQ không cần điền mã hàng, chỉ cần tên Hóa đơn.
    </small>
    """,
    unsafe_allow_html=True
)


# =========================
# DEFAULT TABLE
# =========================

default_products = pd.DataFrame(
    [
        {
            "Tên Sản Phẩm": "",
            "Số lượng": 1,
            "Xuất xứ": "Trung Quốc"
        }
    ]
)


# =========================
# DATA EDITOR
# =========================

product_table = st.data_editor(
    default_products,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,

    column_config={

        "Tên Sản Phẩm":
            st.column_config.TextColumn(
                "Tên Sản Phẩm",
                width="large",
                help=(
                    "Điền tên sản phẩm - "
                    "Ví dụ: Máy tính để bàn "
                    "(ASUS)..."
                )
            ),

        "Số lượng":
            st.column_config.NumberColumn(
                "Số lượng",
                min_value=1,
                step=1,
                default=1,
                width="small"
            ),

        "Xuất xứ":
            st.column_config.SelectboxColumn(
                "Xuất xứ",
                options=[
                    "Trung Quốc",
                    "Đài Loan"
                ],
                default="Trung Quốc",
                width="medium"
            )
    },

    key="product_editor"
)


# =========================
# READ PRODUCT TABLE
# =========================

clean_rows = []

table_records = (
    product_table
    .to_dict("records")
)

for row in table_records:

    product_name = str(
        row.get(
            "Tên Sản Phẩm",
            ""
        )
    ).strip()

    if not product_name:
        continue

    quantity = row.get(
        "Số lượng",
        1
    )

    if pd.isna(quantity):
        quantity = 1

    quantity = int(
        quantity
    )

    origin = row.get(
        "Xuất xứ",
        "Trung Quốc"
    )

    if (
        pd.isna(origin)
        or not str(origin).strip()
    ):
        origin = "Trung Quốc"

    origin = str(
        origin
    ).strip()

    clean_rows.append(
        {
            "product_name":
                product_name,

            "quantity":
                quantity,

            "origin":
                origin
        }
    )


# =========================
# SERIAL NUMBER
# =========================

products = []

if clean_rows:

    st.markdown(
        "### Serial Number"
    )

    st.caption(
        "Paste Serial Number riêng cho từng sản phẩm. "
        "Mỗi dòng 1 Serial. "
        "Có thể copy trực tiếp từ Excel."
    )

    for i, row in enumerate(
        clean_rows
    ):

        st.markdown(
            f"**{i + 1}. "
            f"{row['product_name']}**"
        )

        serial_text = st.text_area(
            (
                f"Serial Number - "
                f"Sản phẩm {i + 1}"
            ),
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
                f"Đã nhận: "
                f"{len(serials)} "
                f"Serial Number"
            )

        products.append(
            {
                "product_name":
                    row[
                        "product_name"
                    ],

                "quantity":
                    row[
                        "quantity"
                    ],

                "origin":
                    row[
                        "origin"
                    ],

                "serials":
                    serials
            }
        )

        st.write("")

else:

    st.info(
        "Hãy nhập ít nhất 1 "
        "Tên Sản Phẩm trong bảng "
        "để nhập Serial Number."
    )


# =========================
# CHECK DUPLICATE SERIAL
# =========================

serial_locations = {}

for product in products:

    for serial in product[
        "serials"
    ]:

        serial_key = (
            serial
            .strip()
            .upper()
        )

        if (
            serial_key
            not in serial_locations
        ):
            serial_locations[
                serial_key
            ] = []

        serial_locations[
            serial_key
        ].append(
            product[
                "product_name"
            ]
        )


duplicate_serials = {
    serial:
        product_names

    for serial, product_names
    in serial_locations.items()

    if len(product_names) > 1
}


if duplicate_serials:

    st.error(
        f"⚠️ Phát hiện "
        f"{len(duplicate_serials)} "
        f"Serial Number bị trùng. "
        f"Vui lòng kiểm tra lại "
        f"trước khi tạo CQ."
    )

    for (
        serial,
        product_names
    ) in duplicate_serials.items():

        st.markdown(
            f"🔴 **{serial}** — "
            f"xuất hiện "
            f"{len(product_names)} lần"
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
        "Có ít nhất 1 sản phẩm "
        "từ 6 Serial Number trở lên. "
        "Toàn bộ Serial Number của CQ "
        "sẽ được đưa xuống phụ lục."
    )


# =========================
# SUMMARY
# =========================

if products:

    total_models = len(
        products
    )

    total_serials = sum(
        len(
            product["serials"]
        )
        for product
        in products
    )

    st.caption(
        f"{total_models} sản phẩm | "
        f"{total_serials} Serial Number"
    )


# =========================
# GENERATE CQ
# =========================

st.divider()


if st.button(
    "GENERATE CQ",
    type="primary",
    use_container_width=True,
    disabled=bool(
        duplicate_serials
    )
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

    for index, product in enumerate(
        products
    ):

        if not product[
            "product_name"
        ].strip():

            errors.append(
                f"Sản phẩm {index + 1}: "
                f"chưa nhập Tên Sản Phẩm."
            )

        if not product[
            "origin"
        ].strip():

            errors.append(
                f"Sản phẩm {index + 1}: "
                f"chưa chọn Xuất xứ."
            )

    if errors:

        for error in errors:

            st.error(
                error
            )

        st.stop()


    # =========================
    # PROCESS SERIAL DATA
    # =========================

    document_products = []

    appendix_products = []

    use_appendix_for_all = any(
        len(
            product["serials"]
        ) >= 6
        for product
        in products
    )

    for product in products:

        serials = product[
            "serials"
        ]

        if use_appendix_for_all:

            serial_display = (
                "Phụ lục đính kèm"
            )

            appendix_products.append(
                {
                    "product_name":
                        product[
                            "product_name"
                        ],

                    "serial_rows":
                        serial_to_rows(
                            serials,
                            columns=4
                        )
                }
            )

        else:

            serial_display = (
                "\n".join(
                    serials
                )
            )

        document_products.append(
            {
                "product_name":
                    product[
                        "product_name"
                    ],

                "quantity":
                    product[
                        "quantity"
                    ],

                "origin":
                    product[
                        "origin"
                    ],

                "serial_display":
                    serial_display
            }
        )


    # =========================
    # CITY TEXT
    # =========================

    city_text = city

    if city == "TP.HCM":

        city_text = "Tp.HCM"


    # =========================
    # SELECT TEMPLATE
    # =========================

    if form_type == "ASUS":

        template_file = (
            "CQ_SYSTEM_TEMPLATE_OPTIONAL_APPENDIX.docx"
        )

        file_prefix = (
            "CQ_ASUS"
        )

    elif (
        form_type == "DGW"
        and city == "Hà Nội"
    ):

        template_file = (
            "CQ_DGW_TEMPLATE_HN_FIXED.docx"
        )

        file_prefix = (
            "CQ_DGW_HN"
        )

    else:

        template_file = (
            "CQ_DGW_TEMPLATE_HCM_FIXED.docx"
        )

        file_prefix = (
            "CQ_DGW_HCM"
        )


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

        if not os.path.exists(
            template_file
        ):

            st.error(
                f"Không tìm thấy template: "
                f"{template_file}"
            )

            st.stop()

        doc = DocxTemplate(
            template_file
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
        # =========================

        safe_eu = re.sub(
            r'[\\/*?:"<>|]',
            "",
            eu_name
        )

        safe_eu = (
            safe_eu.strip()
        )

        filename = (
            f"{file_prefix}_"
            f"{safe_eu}_"
            f"{issue_date.strftime('%Y%m%d')}"
            f".docx"
        )


        # =========================
        # WORD BYTES
        # =========================

        word_bytes = (
            output.getvalue()
        )


        # =========================
        # GOOGLE DRIVE + TRACKING
        # =========================

        try:

            # Drive dùng OAuth
            drive_service = (
                get_drive_service()
            )

            # Sheet dùng Service Account
            sheets_service = (
                get_sheets_service()
            )


            # =========================
            # UPLOAD WORD TO DRIVE
            # =========================

            drive_link = (
                upload_cq_to_drive(
                    drive_service=
                        drive_service,

                    word_bytes=
                        word_bytes,

                    filename=
                        filename,

                    issue_date=
                        issue_date
                )
            )


            # =========================
            # ADD TO CQ TRACKING
            # =========================

            append_tracking_row(
                sheets_service=
                    sheets_service,

                creator_name=
                    creator_name,

                drive_link=
                    drive_link,

                issue_date=
                    issue_date,

                form_type=
                    form_type,

                city=
                    city,

                eu_name=
                    eu_name,

                address=
                    address,

                products=
                    products
            )


            st.success(
                "CQ đã được tạo thành công.\n\n"
                "✅ Đã lưu Word lên Google Drive.\n\n"
                "✅ Đã thêm vào CQ TRACKING."
            )


            # =========================
            # DRIVE LINK
            # =========================

            st.link_button(
                "MỞ CQ TRÊN GOOGLE DRIVE",
                drive_link,
                use_container_width=True
            )


            # =========================
            # TRACKING LINK
            # =========================

            sheet_id = (
                st.secrets[
                    "google"
                ][
                    "sheet_id"
                ]
            )

            tracking_link = (
                "https://docs.google.com/"
                "spreadsheets/d/"
                f"{sheet_id}/edit"
            )

            st.link_button(
                "MỞ CQ TRACKING",
                tracking_link,
                use_container_width=True
            )


        except Exception as google_error:

            st.warning(
                "CQ đã tạo được Word "
                "nhưng chưa lưu được lên "
                "Google Drive / CQ TRACKING."
            )

            st.error(
                f"Google error: "
                f"{str(google_error)}"
            )


        # =========================
        # DOWNLOAD WORD
        # =========================

        st.download_button(
            label="DOWNLOAD WORD",
            data=word_bytes,
            file_name=filename,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True
        )


    except FileNotFoundError:

        st.error(
            f"Không tìm thấy file template: "
            f"{template_file}"
        )

    except Exception as e:

        st.error(
            f"Có lỗi khi tạo CQ: "
            f"{str(e)}"
        )
