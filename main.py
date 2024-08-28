import io

import msoffcrypto
import PyPDF2
import streamlit as st

# ================================
# ヘルパー関数の定義
# ================================


def is_pdf_encrypted(file_bytes):
    """
    PDFファイルが暗号化されているかどうかを判定します。

    Args:
        file_bytes (bytes): PDFファイルのバイトデータ。

    Returns:
        bool: 暗号化されている場合はTrue、そうでない場合はFalse。
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        return reader.is_encrypted
    except Exception:
        return False


def decrypt_pdf(file_bytes, password):
    """
    PDFファイルを指定されたパスワードで復号します。

    Args:
        file_bytes (bytes): PDFファイルのバイトデータ。
        password (str): 復号に使用するパスワード。

    Returns:
        bool: 復号に成功した場合はTrue、失敗した場合はFalse。
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        result = reader.decrypt(password)
        return result != 0  # 0は失敗、1または2は成功
    except Exception:
        return False


def is_office_encrypted(file_bytes):
    """
    Officeファイル（DOCX、XLSX、PPTX）が暗号化されているかどうかを判定します。

    Args:
        file_bytes (bytes): Officeファイルのバイトデータ。

    Returns:
        bool: 暗号化されている場合はTrue、そうでない場合はFalse。
    """
    try:
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        return office_file.is_encrypted()
    except Exception:
        return False


def decrypt_office(file_bytes, password):
    """
    Officeファイルを指定されたパスワードで復号します。

    Args:
        file_bytes (bytes): Officeファイルのバイトデータ。
        password (str): 復号に使用するパスワード。

    Returns:
        bytes or None: 復号に成功した場合は復号後のバイトデータ、失敗した場合はNone。
    """
    try:
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        return decrypted.getvalue()
    except Exception:
        return None


def get_current_file():
    """
    現在処理中の暗号化されたファイルを取得します。

    Returns:
        str or None: 現在処理中のファイル名、なければNone。
    """
    encrypted_files = [
        f
        for f in st.session_state["file_status"]
        if st.session_state["file_status"][f]["encrypted"]
        and not st.session_state["file_status"][f]["decrypted"]
        and not st.session_state["file_status"][f]["excluded"]
    ]
    if encrypted_files:
        return encrypted_files[0]
    else:
        return None


# ================================
# コールバック関数の定義
# ================================


def handle_submit():
    """
    パスワードの「Submit」ボタンが押されたときの処理。
    """
    current_file = get_current_file()
    if not current_file:
        return

    password_attempt = st.session_state.get("password_input", "")
    if not password_attempt:
        st.session_state["error_message"] = "パスワードを入力してください。"
        return

    extension = st.session_state["file_status"][current_file]["extension"]
    file_bytes = st.session_state["file_status"][current_file]["file_bytes"]

    if extension == "pdf":
        success = decrypt_pdf(file_bytes, password_attempt)
    else:  # docx, xlsx, pptx
        decrypted_bytes = decrypt_office(file_bytes, password_attempt)
        success = decrypted_bytes is not None
        if success:
            st.session_state["file_status"][current_file]["content"] = decrypted_bytes

    if success:
        st.session_state["file_status"][current_file]["decrypted"] = True
        st.session_state["decrypted_files"].append(current_file)
        st.success(f"ファイル '{current_file}' を復号しました。")
        st.session_state["error_message"] = ""
        # パスワード入力をリセット
        st.session_state["password_input"] = ""
    else:
        st.session_state["error_message"] = "パスワードが正しくありません。再度お試しください。"


def handle_cancel():
    """
    パスワードの「Cancel」ボタンが押されたときの処理。
    """
    current_file = get_current_file()
    if not current_file:
        return

    st.session_state["file_status"][current_file]["excluded"] = True
    st.session_state["excluded_files"].append(current_file)
    st.session_state["error_message"] = f"ファイル '{current_file}' の処理をキャンセルしました。"
    # パスワード入力をリセット
    st.session_state["password_input"] = ""


@st.dialog("パスワードを入力してください。")
def password_dialog():
    """パスワード入力ダイアログの UI コンポーネント"""
    current_file = get_current_file()

    if not current_file:
        st.rerun()

    st.markdown(f"### ファイル '{current_file}' は暗号化されています。パスワードを入力してください。")

    # エラーメッセージの表示
    if st.session_state["error_message"]:
        st.error(st.session_state["error_message"])

    # パスワード入力とボタン
    password = st.text_input("パスワード", type="password", key="password_input")

    col1, col2 = st.columns(2)

    with col1:
        st.button(
            "Submit", key="submit_password", on_click=handle_submit, type="primary"
        )

    with col2:
        st.button("Cancel", key="cancel_password", on_click=handle_cancel)


# ================================
# セッションステートの初期化
# ================================

if "file_status" not in st.session_state:
    st.session_state["file_status"] = {}

if "processing" not in st.session_state:
    st.session_state["processing"] = False

if "decrypted_files" not in st.session_state:
    st.session_state["decrypted_files"] = []

if "excluded_files" not in st.session_state:
    st.session_state["excluded_files"] = []

if "error_message" not in st.session_state:
    st.session_state["error_message"] = ""

if "password_input" not in st.session_state:
    st.session_state["password_input"] = ""

# ================================
# Streamlitアプリケーションの本体
# ================================

st.title("ファイル復号アプリケーション")

# ファイルアップローダー
uploaded_files = st.file_uploader(
    "PDF, DOCX, XLSX, PPTXファイルをアップロードしてください。",
    type=["pdf", "docx", "xlsx", "pptx"],
    accept_multiple_files=True,
)

# 「実行」ボタン
if st.button("実行"):
    if uploaded_files:
        # セッションステートをリセット
        st.session_state["processing"] = True
        st.session_state["file_status"] = {}
        st.session_state["decrypted_files"] = []
        st.session_state["excluded_files"] = []
        st.session_state["error_message"] = ""
        st.session_state["password_input"] = ""

        # アップロードされた各ファイルを処理
        for file in uploaded_files:
            file_bytes = file.read()
            st.session_state["file_status"][file.name] = {
                "extension": file.name.split(".")[-1].lower(),
                "encrypted": False,
                "decrypted": False,
                "excluded": False,
                "content": None,
                "file_bytes": file_bytes,
                "error": None,
            }

            extension = st.session_state["file_status"][file.name]["extension"]

            if extension == "pdf":
                encrypted = is_pdf_encrypted(file_bytes)
            elif extension in ["docx", "xlsx", "pptx"]:
                encrypted = is_office_encrypted(file_bytes)
            else:
                # サポートされていない拡張子の場合は暗号化されていないと見なす
                encrypted = False

            st.session_state["file_status"][file.name]["encrypted"] = encrypted

            if not encrypted:
                st.session_state["file_status"][file.name]["decrypted"] = True
                st.session_state["decrypted_files"].append(file.name)
    else:
        st.warning("少なくとも1つのファイルをアップロードしてください。")

# 処理中の場合
if st.session_state["processing"]:
    current_file = get_current_file()

    if current_file:
        password_dialog()
    else:
        # すべての暗号化されたファイルが処理された場合
        st.session_state["processing"] = False
        st.success("すべてのファイルの処理が完了しました。")

        if st.session_state["decrypted_files"]:
            st.markdown("### 読み取ることができたファイルの名前:")
            for f in st.session_state["decrypted_files"]:
                st.write(f)
        else:
            st.info("読み取ることができたファイルはありませんでした。")
        if st.session_state["excluded_files"]:
            st.markdown("### 読み取ることができなかったファイルの名前:")
            for f in st.session_state["excluded_files"]:
                st.write(f)
