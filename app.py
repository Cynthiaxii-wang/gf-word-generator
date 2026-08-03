from pathlib import Path
import streamlit as st
import subprocess
import tempfile
import shutil
import sys


# =====================================================
# Path
# =====================================================

ROOT_DIR = Path(__file__).resolve().parent


PIPELINE_PATH = (
    ROOT_DIR
    /
    "scripts"
    /
    "run_pipeline.py"
)


TEMPLATE_PATH = (
    ROOT_DIR
    /
    "template"
    /
    "总量通用.docx"
)



# =====================================================
# Streamlit config
# =====================================================

st.set_page_config(
    page_title="广发策略报告生成器",
    page_icon="📄",
    layout="centered"
)


st.title("📄 广发策略报告生成器")


st.markdown(
    """
上传研究员原始 Word 文件，
自动生成符合广发策略模板格式的研究报告。

支持：

- Word 模板保留
- 原生 Chart 对象复制
- 表格格式复制
- 自动标题识别
"""
)



# =====================================================
# Upload
# =====================================================

uploaded_file = st.file_uploader(
    "请选择 Word 文件 (.docx)",
    type=["docx"]
)



if uploaded_file:


    st.success(
        f"已上传：{uploaded_file.name}"
    )



    if st.button(
        "🚀 开始生成",
        type="primary"
    ):


        with st.spinner(
            "正在解析文档、复制原生对象、生成报告..."
        ):


            try:


                # ---------------------------------
                # 创建独立临时目录
                # ---------------------------------

                with tempfile.TemporaryDirectory() as tmp:


                    tmp = Path(tmp)


                    input_dir = (
                        tmp
                        /
                        "input"
                    )


                    work_dir = (
                        tmp
                        /
                        "work"
                    )


                    output_dir = (
                        tmp
                        /
                        "output"
                    )


                    input_dir.mkdir()
                    work_dir.mkdir()
                    output_dir.mkdir()



                    # ---------------------------------
                    # 保存用户上传文件
                    # ---------------------------------

                    input_path = (
                        input_dir
                        /
                        uploaded_file.name
                    )


                    with open(
                        input_path,
                        "wb"
                    ) as f:

                        f.write(
                            uploaded_file.getbuffer()
                        )



                    # ---------------------------------
                    # 调用 pipeline
                    # ---------------------------------

                    command = [

                        sys.executable,

                        str(PIPELINE_PATH),

                        str(input_path),

                        "--work-dir",

                        str(work_dir),

                        "--output-dir",

                        str(output_dir),

                        "--template",

                        str(TEMPLATE_PATH)

                    ]



                    result = subprocess.run(

                        command,

                        cwd=str(ROOT_DIR),

                        capture_output=True,

                        text=True

                    )



                    # ---------------------------------
                    # pipeline失败
                    # ---------------------------------

                    if result.returncode != 0:


                        st.error(
                            "生成失败"
                        )


                        st.code(
                            result.stderr
                        )



                    else:


                        st.success(
                            "报告生成完成！"
                        )


                        # -----------------------------
                        # 找生成文件
                        # -----------------------------

                        docx_files = list(
                            output_dir.glob(
                                "*.docx"
                            )
                        )


                        if not docx_files:


                            st.warning(
                                "未找到生成文件"
                            )


                            st.text(
                                result.stdout
                            )


                        else:


                            output_file = max(

                                docx_files,

                                key=lambda x:
                                x.stat().st_mtime

                            )


                            st.info(
                                output_file.name
                            )



                            # -----------------------------
                            # 下载
                            # -----------------------------

                            with open(
                                output_file,
                                "rb"
                            ) as f:


                                st.download_button(

                                    label="⬇️ 下载报告",

                                    data=f,

                                    file_name=output_file.name,

                                    mime=
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                                )



            except Exception as e:


                st.error(
                    f"程序异常：{e}"
                )