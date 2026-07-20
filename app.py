"""
体检报告在线生成器 — Streamlit Web App
在浏览器中填写信息，下载生成的 Word 体检报告
"""
import os, sys, shutil, re, tempfile, io
from datetime import date, datetime

import streamlit as st
import openpyxl

from tihuan_tijian import (
    find_target, step2_replace,
    _find_idcard, _find_phone, _find_date, _find_time, _find_gender,
    replace_gender_in_doc,
    CUSTOM_ITEMS, ALL_ITEMS
)

# ─── 路径 ─────────────────────────────────────────────
WORK_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(WORK_DIR, "王显勇.docx")


# ─── 工具函数 ─────────────────────────────────────────
def calc_age(idcard):
    """从身份证号计算年龄"""
    m = re.match(r'\d{6}(\d{4})(\d{2})(\d{2})\d{3}[\dXx]', str(idcard))
    if not m:
        return ''
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    today = date.today()
    age = today.year - y
    if (today.month, today.day) < (mo, d):
        age -= 1
    return str(age)


def detect_gender(idcard):
    """从身份证号判断性别：17位奇男偶女"""
    m = re.match(r'\d{17}[\dXx]', str(idcard))
    if not m:
        return None
    digit = str(idcard)[16]
    if digit in 'Xx':
        return None
    return '女' if int(digit) % 2 == 0 else '男'


def generate_report(person):
    """生成体检报告，返回 docx 文件路径"""
    out = os.path.join(WORK_DIR, f"_temp_{person.get('name','temp')}.docx")
    shutil.copy(TEMPLATE, out)

    import docx
    doc = docx.Document(out)

    # 构建替换数据
    new_vals = {}
    for item in ALL_ITEMS:
        val = person.get(item.key)
        if val:
            new_vals[item.key] = str(val)

    # 性别替换
    gender = person.get('gender')
    if gender:
        old_gen, _, _ = _find_gender(doc)
        if old_gen and old_gen != gender:
            count = replace_gender_in_doc(doc, old_gen, gender)
            if count > 0:
                backup = out.replace('.docx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
                shutil.copy(out, backup)
                doc.save(out)
                doc = docx.Document(out)

    # 标准替换
    key_to_item = {it.key: it for it in ALL_ITEMS}
    replacements = []
    for key, new_val in new_vals.items():
        if key == 'gender':
            continue
        item = key_to_item.get(key)
        if item is None:
            continue
        old_val, _, _ = find_target(doc, item)
        if old_val is None:
            continue
        if key == 'age':
            replacements.append((item.label, old_val + '岁', new_val + '岁'))
            # 处理 "5 2 岁" 这种带空格的年龄（2位以上数字）
            if len(old_val) >= 2 and len(new_val) >= 2:
                replacements.append((item.label,
                    old_val[0] + ' ' + old_val[1] + ' ' + '岁',
                    new_val[0] + ' ' + new_val[1] + ' ' + '岁'))
        elif key == 'height':
            replacements.append((item.label, old_val + 'cm', new_val + 'cm'))
        elif key == 'weight':
            replacements.append((item.label, old_val + 'Kg', new_val + 'Kg'))
        elif key == 'sbp':
            replacements.append((item.label, old_val + 'mmHg', new_val + 'mmHg'))
        elif key == 'dbp':
            replacements.append((item.label, old_val + 'mmHg', new_val + 'mmHg'))
            replacements.append(('心率', old_val + 'bpm', new_val + 'bpm'))
        else:
            replacements.append((item.label, old_val, new_val))

    if replacements:
        step2_replace(out, replacements=replacements)

    return out


def convert_docx_to_pdf(docx_path):
    """使用 LibreOffice 将 docx 转为 pdf，返回 pdf 路径"""
    import subprocess, os, random
    pdf_path = docx_path.replace('.docx', '.pdf')
    tmp_profile = f"/tmp/lo_{random.randint(10000,99999)}"
    cmd = [
        'soffice', '--headless', '--convert-to', 'pdf',
        '--outdir', os.path.dirname(docx_path),
        f'-env:UserInstallation=file://{tmp_profile}',
        docx_path
    ]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        return None
    return pdf_path if os.path.exists(pdf_path) else None


# ─── 页面 ─────────────────────────────────────────────
st.set_page_config(
    page_title="体检报告生成器",
    page_icon="📋",
    layout="centered"
)

st.title("📋 体检报告生成器")

mode = st.radio("选择模式", ["📝 单人生成", "📦 批量生成（上传 Excel）"], horizontal=True)

if mode == "📦 批量生成（上传 Excel）":
    st.markdown("上传 Excel 文件（格式同 `人员信息模板.xlsx`），批量生成所有人的体检报告。")
    uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"])

    if uploaded:
        wb = openpyxl.load_workbook(io.BytesIO(uploaded.read()))
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            st.error("Excel 中没有数据（至少需要表头+一行数据）")
            st.stop()

        headers = [str(h).strip() if h else '' for h in rows[0]]
        col_map = {
            '姓名': 'name', '身份证号': 'idcard', '体检编码': 'number',
            '身高': 'height', '体重': 'weight', '体重指数': 'bmi',
            '收缩压': 'sbp', '舒张压': 'dbp', '电话': 'phone',
            '体检日期': 'date', '检查时间': 'time',
        }

        persons = []
        for row in rows[1:]:
            if not row or all(v is None or str(v).strip() == '' for v in row):
                continue
            p = {}
            for hi, h in enumerate(headers):
                key = col_map.get(h)
                if key and hi < len(row) and row[hi] is not None:
                    val = str(row[hi]).strip()
                    if val:
                        p[key] = val
            if 'name' not in p:
                continue
            # 自动计算年龄和性别
            if 'idcard' in p:
                p['age'] = calc_age(p['idcard'])
                g = detect_gender(p['idcard'])
                if g: p['gender'] = g
            # 格式化数值
            for k in ('height', 'weight', 'bmi'):
                if k in p:
                    try: p[k] = f"{float(p[k]):.2f}"
                    except: pass
            persons.append(p)

        if not persons:
            st.error("未能解析到有效的人员数据")
            st.stop()

        st.success(f"共解析到 {len(persons)} 人：{', '.join(p.get('name','') for p in persons)}")

        # 确认按钮
        if st.button("🚀 批量生成全部报告", use_container_width=True):
            import zipfile
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                prog = st.progress(0, text="正在生成...")
                for idx, p in enumerate(persons):
                    name = p.get('name', f'person_{idx+1}')
                    prog.progress((idx) / len(persons), text=f"正在生成 {name}...")
                    try:
                        path = generate_report(p)
                        # docx
                        with open(path, 'rb') as f:
                            zf.writestr(f"{name}.docx", f.read())
                        # 尝试转 PDF
                        try:
                            pdf_path = convert_docx_to_pdf(path)
                            if pdf_path:
                                with open(pdf_path, 'rb') as f:
                                    zf.writestr(f"{name}.pdf", f.read())
                                os.remove(pdf_path)
                        except:
                            pass
                        os.remove(path)
                    except Exception as e:
                        st.error(f"{name} 生成失败: {e}")
                prog.progress(1.0, text="完成！")

            st.success(f"✅ 已生成 {len(persons)} 份体检报告！（含 PDF）")
            st.download_button(
                label="📥 下载全部报告（含 Word + PDF）",
                data=buf.getvalue(),
                file_name="体检报告_批量.zip",
                mime="application/zip",
                use_container_width=True,
            )

elif mode == "📝 单人生成":
    st.markdown("填写下方信息（全部为必填），自动生成 Word 体检报告文件。")

with st.form("report_form"):
    col1, col2 = st.columns(2)

    with col1:
        name = st.text_input("姓名 *", placeholder="请输入姓名")
        idcard = st.text_input("身份证号 *", max_chars=18, placeholder="18 位身份证号",
                               help="年龄从身份证号自动计算")

    with col2:
        gender = st.radio("性别 *", ["男", "女"], horizontal=True,
                          help="如填写身份证号，性别会自动匹配")
        exam_code = st.text_input("体检编码 *", placeholder="如 202607080007")

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        height = st.number_input("身高 (cm) *", min_value=100, max_value=250, value=165, step=1)
    with col_b:
        weight = st.number_input("体重 (Kg) *", min_value=20, max_value=200, value=65, step=1)
    with col_c:
        bmi = st.number_input("BMI *", min_value=10.0, max_value=50.0, value=22.5, format="%.2f")

    col_d, col_e, col_f = st.columns(3)
    with col_d:
        sbp = st.number_input("收缩压 (mmHg) *", min_value=80, max_value=220, value=120, step=1)
    with col_e:
        dbp = st.number_input("舒张压 (mmHg) *", min_value=40, max_value=140, value=80, step=1)
    with col_f:
        phone = st.text_input("联系电话 *", max_chars=11, placeholder="11 位手机号")

    col_g, col_h = st.columns(2)
    with col_g:
        exam_date = st.date_input("体检日期 *", value=date.today())
    with col_h:
        exam_time = st.time_input("检查时间 *", value=datetime.now().time())

    submitted = st.form_submit_button("🚀 生成体检报告", use_container_width=True)

if submitted:
    # 所有字段必填校验
    errors = []
    if not name or not name.strip():
        errors.append("姓名")
    if not idcard or len(idcard.strip()) != 18:
        errors.append("身份证号（需 18 位）")
    if not exam_code or not exam_code.strip():
        errors.append("体检编码")
    if not phone or len(phone.strip()) < 11:
        errors.append("联系电话")
    # 数值字段默认已有值，无需校验

    if errors:
        st.error(f"⚠️ 请填写以下必填项：{'、'.join(errors)}")
        st.stop()

    idcard = idcard.strip()

    # 身份证号性别校验
    detected_gender = detect_gender(idcard)
    if detected_gender and detected_gender != gender:
        st.warning(f"⚠️ 选的性别是「{gender}」，但身份证号第 17 位显示为「{detected_gender}」，请确认")

    person = {
        'name': name.strip(),
        'idcard': idcard,
        'number': exam_code.strip(),
        'gender': gender,
        'height': str(height),
        'weight': str(weight),
        'bmi': f"{bmi:.2f}",
        'sbp': str(sbp),
        'dbp': str(dbp),
        'phone': phone.strip(),
        'date': exam_date.strftime("%Y-%m-%d"),
        'time': exam_time.strftime("%H:%M"),
        'age': calc_age(idcard),
    }

    with st.spinner(f"正在生成 {person['name']} 的体检报告..."):
        try:
            out_path = generate_report(person)

            # 尝试转 PDF
            pdf_path = None
            try:
                pdf_path = convert_docx_to_pdf(out_path)
            except:
                pass

            with open(out_path, "rb") as f:
                docx_bytes = f.read()

            pdf_bytes = None
            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                os.remove(pdf_path)
            os.remove(out_path)

            st.success(f"✅ **{person['name']}** 的体检报告已生成！")

            gender_text = person.get('gender', '男')
            st.info(
                f"📄 **{person['name']}** | {'♂' if gender_text == '男' else '♀'} {gender_text} | "
                f"{person['age']}岁 | 体检编码: {person['number']}"
            )

            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button(
                    label="📥 下载 Word",
                    data=docx_bytes,
                    file_name=f"{person['name']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            if pdf_bytes:
                with col_b:
                    st.download_button(
                        label="📄 下载 PDF",
                        data=pdf_bytes,
                        file_name=f"{person['name']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

        except Exception as e:
            st.error(f"❌ 生成失败: {e}")
            import traceback
            st.exception(e)

else:
    st.markdown("---")
    st.caption("全部字段均为必填。年龄从身份证号自动计算。")
