import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests
import streamlit as st


st.set_page_config(
    page_title="Atendimento Facilitado",
    page_icon="🧾",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stSidebar"] {display: none !important;}
    [data-testid="collapsedControl"] {display: none !important;}

    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2rem;
        max-width: 780px;
    }

    .title-wrap {
        text-align: center;
        margin-bottom: 1.2rem;
    }

    .tiny {
        color: #94a3b8;
        font-size: 0.95rem;
        margin-top: 0.25rem;
    }

    .card {
        padding: 1.2rem 1.2rem 0.8rem 1.2rem;
        border: 1px solid rgba(148,163,184,.15);
        border-radius: 18px;
        background: rgba(255,255,255,.03);
        margin-bottom: 1rem;
    }

    .bottom-box {
        margin-top: 2rem;
        padding-top: 1rem;
        border-top: 1px solid rgba(148,163,184,.18);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

FIELDS = [
    {
        "key": "pessoa_solicitante",
        "label": "Pessoa solicitante",
        "type": "text",
        "help": "Pessoa responsável pela abertura do atendimento. Ex.: lojista, cliente, esposa, filho.",
    },
    {
        "key": "modelo",
        "label": "Modelo",
        "type": "text",
        "help": "Modelo do produto a ser reparado.",
    },
    {
        "key": "ns",
        "label": "NS",
        "type": "text",
        "help": "Número de série do produto.",
    },
    {
        "key": "fato",
        "label": "Fato",
        "type": "textarea",
        "help": "Descrição do problema apresentado pelo produto.",
    },
    {
        "key": "causa",
        "label": "Causa",
        "type": "textarea",
        "help": "Causa identificada para o problema.",
    },
    {
        "key": "acao",
        "label": "Ação",
        "type": "textarea",
        "help": "Ação realizada para resolução do problema.",
    },
    {
        "key": "cep",
        "label": "CEP",
        "type": "cep",
        "help": "CEP do endereço para retorno do produto.",
    },
    {
        "key": "rua",
        "label": "Rua",
        "type": "text",
        "help": "Será preenchido automaticamente quando o CEP for encontrado.",
    },
    {
        "key": "numero",
        "label": "Número",
        "type": "text",
        "help": "Número do local de devolução.",
    },
    {
        "key": "complemento",
        "label": "Complemento",
        "type": "text",
        "help": "Informação adicional do endereço.",
    },
    {
        "key": "bairro",
        "label": "Bairro",
        "type": "text",
        "help": "Será preenchido automaticamente quando o CEP for encontrado.",
    },
    {
        "key": "cidade",
        "label": "Cidade",
        "type": "text",
        "help": "Será preenchido automaticamente quando o CEP for encontrado.",
    },
    {
        "key": "estado",
        "label": "Estado",
        "type": "text",
        "help": "UF. Ex.: SP, RJ, MG.",
    },
]


def init_state():
    if "step" not in st.session_state:
        st.session_state.step = 0

    if "form_data" not in st.session_state:
        st.session_state.form_data = {field["key"]: "" for field in FIELDS}

    if "session_id" not in st.session_state:
        st.session_state.session_id = datetime.now().strftime("%Y%m%d%H%M%S")

    if "address_locked" not in st.session_state:
        st.session_state.address_locked = False

    if "last_imported_file_id" not in st.session_state:
        st.session_state.last_imported_file_id = None


def normalize_cep(cep: str) -> str:
    return re.sub(r"\D", "", cep or "")


def format_cep(cep: str) -> str:
    cep = normalize_cep(cep)
    if len(cep) == 8:
        return f"{cep[:5]}-{cep[5:]}"
    return cep


def is_generic_cep(cep: str) -> bool:
    cep = normalize_cep(cep)
    if len(cep) != 8:
        return True

    invalids = {
        "00000000",
        "11111111",
        "22222222",
        "33333333",
        "44444444",
        "55555555",
        "66666666",
        "77777777",
        "88888888",
        "99999999",
        "12345678",
        "87654321",
    }
    return cep in invalids


def validate_cep_via_viacep(cep: str) -> dict:
    cep_digits = normalize_cep(cep)

    if len(cep_digits) != 8:
        raise ValueError("O CEP precisa ter 8 dígitos.")

    if is_generic_cep(cep_digits):
        raise ValueError("CEP genérico ou inválido.")

    url = f"https://viacep.com.br/ws/{cep_digits}/json/"
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    data = response.json()

    if data.get("erro"):
        raise ValueError("CEP não encontrado.")

    logradouro = (data.get("logradouro") or "").strip()
    bairro = (data.get("bairro") or "").strip()
    cidade = (data.get("localidade") or "").strip()
    estado = (data.get("uf") or "").strip()

    if not cidade or not estado:
        raise ValueError("CEP retornou dados incompletos.")

    return {
        "cep": format_cep(cep_digits),
        "rua": logradouro,
        "bairro": bairro,
        "cidade": cidade,
        "estado": estado,
    }


def auto_fill_address(address: dict):
    for key in ["cep", "rua", "bairro", "cidade", "estado"]:
        if key in address and address[key]:
            st.session_state.form_data[key] = address[key]

    st.session_state.address_locked = True


def build_partial_meta():
    current_step = st.session_state.step
    current_field = FIELDS[min(current_step, len(FIELDS) - 1)]["key"]

    return {
        "session_id": st.session_state.session_id,
        "saved_at": datetime.now().isoformat(),
        "stopped_at_step": current_step,
        "stopped_at_field": current_field,
        "status": "parcial",
    }


def build_final_meta():
    return {
        "session_id": st.session_state.session_id,
        "saved_at": datetime.now().isoformat(),
        "stopped_at_step": len(FIELDS),
        "stopped_at_field": "finalizado",
        "status": "concluido",
    }


def export_excel_bytes(form_data: dict, meta: dict, data_sheet_name: str) -> bytes:
    output = BytesIO()

    df_form = pd.DataFrame([form_data])
    df_meta = pd.DataFrame([meta])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_form.to_excel(writer, index=False, sheet_name=data_sheet_name)
        df_meta.to_excel(writer, index=False, sheet_name="meta")

    output.seek(0)
    return output.getvalue()


def load_partial_excel(uploaded_file):
    df_data = pd.read_excel(uploaded_file, sheet_name="dados")
    df_meta = pd.read_excel(uploaded_file, sheet_name="meta")

    if df_data.empty:
        raise ValueError("A aba 'dados' está vazia.")

    row = df_data.iloc[0].to_dict()
    form_data = {}

    for field in FIELDS:
        key = field["key"]
        value = row.get(key, "")
        form_data[key] = "" if pd.isna(value) else str(value)

    meta_row = {}
    if not df_meta.empty:
        meta_row = df_meta.iloc[0].to_dict()

    meta = {}
    for key, value in meta_row.items():
        meta[str(key)] = "" if pd.isna(value) else value

    if "stopped_at_step" in meta:
        try:
            meta["stopped_at_step"] = int(meta["stopped_at_step"])
        except Exception:
            meta["stopped_at_step"] = 0

    return form_data, meta


def render_progress():
    total = len(FIELDS)
    step = st.session_state.step
    current = min(step + 1, total)
    progress = step / total if total else 0
    st.progress(progress, text=f"Etapa {current}/{total}")


def render_input(field: dict):
    key = field["key"]
    current_value = st.session_state.form_data.get(key, "")

    if field["type"] == "textarea":
        value = st.text_area(
            "",
            value=current_value,
            height=140,
            label_visibility="collapsed",
            placeholder=f"Digite {field['label'].lower()}...",
        )
    elif field["type"] == "cep":
        value = st.text_input(
            "",
            value=current_value,
            label_visibility="collapsed",
            placeholder="00000-000",
            max_chars=9,
        )
    else:
        disabled = st.session_state.address_locked and key in {"rua", "bairro", "cidade", "estado"}
        value = st.text_input(
            "",
            value=current_value,
            label_visibility="collapsed",
            placeholder=f"Digite {field['label'].lower()}...",
            disabled=disabled,
        )

    st.session_state.form_data[key] = value


def render_step():
    step = st.session_state.step
    total = len(FIELDS)

    if step >= total:
        render_review_and_export()
        return

    field = FIELDS[step]

    render_progress()

    st.markdown(
        f"""
        <div class="card">
            <h2 style="margin-bottom:0.2rem;">{field["label"]}</h2>
            <p class="tiny">{field["help"]}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_input(field)

    if st.session_state.address_locked and field["key"] in {"rua", "bairro", "cidade", "estado"}:
        st.caption("Campo preenchido automaticamente a partir do CEP.")

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("← Voltar", use_container_width=True, disabled=step == 0):
            st.session_state.step -= 1
            st.rerun()

    with c2:
        next_label = "Concluir" if step == total - 1 else "Próximo →"
        if st.button(next_label, type="primary", use_container_width=True):
            if field["key"] == "cep":
                cep_value = st.session_state.form_data.get("cep", "")
                if normalize_cep(cep_value):
                    try:
                        result = validate_cep_via_viacep(cep_value)
                        auto_fill_address(result)
                        st.success("CEP validado e endereço preenchido.")
                    except Exception as e:
                        st.warning(f"CEP não validado agora: {e}")

            if field["key"] == "estado":
                st.session_state.form_data["estado"] = (
                    st.session_state.form_data["estado"].strip().upper()
                )

            st.session_state.step += 1
            st.rerun()

    with c3:
        partial_meta = build_partial_meta()
        partial_bytes = export_excel_bytes(
            st.session_state.form_data,
            partial_meta,
            data_sheet_name="dados",
        )

        st.download_button(
            "Salvar e parar",
            data=partial_bytes,
            file_name=f"atendimento_parcial_{st.session_state.session_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def render_review_and_export():
    st.success("Cadastro concluído.")
    st.subheader("Revisão final")

    review_data = st.session_state.form_data.copy()
    review_df = pd.DataFrame([review_data])
    st.dataframe(review_df, use_container_width=True, hide_index=True)

    st.markdown("### Exportação")

    final_meta = build_final_meta()
    final_bytes = export_excel_bytes(
        review_data,
        final_meta,
        data_sheet_name="atendimento",
    )

    st.download_button(
        "Baixar Excel final",
        data=final_bytes,
        file_name=f"atendimento_final_{st.session_state.session_id}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Editar informações", use_container_width=True):
            st.session_state.step = 0
            st.rerun()

    with c2:
        if st.button("Novo atendimento", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()


def render_resume_area():
    st.markdown('<div class="bottom-box">', unsafe_allow_html=True)
    st.markdown("### Retomar atendimento")

    uploaded_file = st.file_uploader(
        "Importe o Excel salvo para continuar de onde parou",
        type=["xlsx"],
        label_visibility="collapsed",
        key="resume_uploader",
    )

    if uploaded_file is not None:
        current_file_id = (
            f"{uploaded_file.name}-"
            f"{uploaded_file.size}-"
            f"{getattr(uploaded_file, 'file_id', 'noid')}"
        )

        if st.session_state.last_imported_file_id != current_file_id:
            try:
                form_data, meta = load_partial_excel(uploaded_file)

                for field in FIELDS:
                    if field["key"] not in form_data:
                        form_data[field["key"]] = ""

                st.session_state.form_data = form_data
                st.session_state.step = meta.get("stopped_at_step", 0)
                st.session_state.session_id = str(
                    meta.get("session_id", datetime.now().strftime("%Y%m%d%H%M%S"))
                )

                cep = st.session_state.form_data.get("cep", "")
                rua = st.session_state.form_data.get("rua", "")
                bairro = st.session_state.form_data.get("bairro", "")
                cidade = st.session_state.form_data.get("cidade", "")
                estado = st.session_state.form_data.get("estado", "")

                st.session_state.address_locked = bool(
                    normalize_cep(cep) and (rua or bairro or cidade or estado)
                )

                st.session_state.last_imported_file_id = current_file_id

                st.success(f"Atendimento restaurado na etapa {st.session_state.step + 1}.")
                st.rerun()

            except Exception as e:
                st.error(f"Não foi possível importar o arquivo: {e}")

    st.markdown("</div>", unsafe_allow_html=True)


init_state()

st.markdown(
    """
    <div class="title-wrap">
        <h1>🧾 Atendimento Facilitado</h1>
        <p class="tiny">Fluxo guiado, etapa por etapa.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

render_step()
render_resume_area()
