import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Email Marketing DB | Odoo",
    page_icon="📧",
    layout="wide"
)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def read_file(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ["utf-8-sig", "utf-8", "latin-1"]:
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("No se pudo leer el CSV con las codificaciones estándar.")
    return pd.read_excel(file)


def norm_key(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    return str(val).strip().lower()


def norm_email(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    return str(val).strip().lower()


def clean_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def validate_columns(df, required_cols, label):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(
            f"**{label}**: faltan las columnas `{missing}`. "
            f"Columnas encontradas: `{df.columns.tolist()}`"
        )
        return False
    return True


# ─── Columnas esperadas ────────────────────────────────────────────────────────

B1_COLS = ["Fecha creación", "Cliente"]
B2_COLS = ["Nombre", "Nombre mostrado", "Correo electrónico"]
B3_COLS = ["Nombre", "Compañía relacionada", "Correo electrónico"]


# ─── UI ───────────────────────────────────────────────────────────────────────

st.title("📧 Generador de Base de Datos Email Marketing")
st.caption("Sube las 3 bases exportadas desde Odoo para generar tu lista de email marketing.")

st.subheader("1. Base de Pedidos Odoo")
st.caption("Filtro en Odoo: part number de la marca + estado = pedido de venta")
st.caption("Columnas requeridas: `Fecha creación`, `Cliente`")
file1 = st.file_uploader("Subir Base 1", type=["xlsx", "csv"], key="b1")

st.divider()

st.subheader("2. Base de Empresas")
st.caption("Filtro en Odoo: Compañía relacionada - No está establecida")
st.caption("Columnas requeridas: `Nombre`, `Nombre mostrado`, `Correo electrónico`")
file2 = st.file_uploader("Subir Base 2", type=["xlsx", "csv"], key="b2")

st.divider()

st.subheader("3. Base de Contactos Internos")
st.caption("Filtro en Odoo: Compañía relacionada - Está establecida")
st.caption("Columnas requeridas: `Nombre`, `Compañía relacionada`, `Correo electrónico`")
file3 = st.file_uploader("Subir Base 3", type=["xlsx", "csv"], key="b3")

st.divider()


# ─── Cargar y validar ─────────────────────────────────────────────────────────

df1 = df2 = df3 = None
b1_ok = b2_ok = b3_ok = False

if file1:
    try:
        df1 = read_file(file1)
        b1_ok = validate_columns(df1, B1_COLS, "Base 1")
    except Exception as e:
        st.error(f"Error al leer Base 1: {e}")

if file2:
    try:
        df2 = read_file(file2)
        b2_ok = validate_columns(df2, B2_COLS, "Base 2")
    except Exception as e:
        st.error(f"Error al leer Base 2: {e}")

if file3:
    try:
        df3 = read_file(file3)
        b3_ok = validate_columns(df3, B3_COLS, "Base 3")
    except Exception as e:
        st.error(f"Error al leer Base 3: {e}")

all_ready = b1_ok and b2_ok and b3_ok

if not all_ready:
    st.info("Sube las 3 bases con las columnas correctas para habilitar el procesamiento.")
else:
    if st.button("🚀 Procesar", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):

            # ── 1. Empresas únicas de Base 1 ──────────────────────────────────
            companies_keys = {
                k
                for k in (norm_key(v) for v in df1["Cliente"].dropna())
                if k is not None
            }
            display_map = {}
            for v in df1["Cliente"].dropna():
                k = norm_key(v)
                if k and k not in display_map:
                    display_map[k] = clean_str(v)
            total_companies = len(companies_keys)

            # ── 2. Filtrar Base 2 a empresas de Base 1 ────────────────────────
            df2 = df2.copy()
            df2["_key"] = df2["Nombre mostrado"].apply(norm_key)
            df2["_email"] = df2["Correo electrónico"].apply(norm_email)
            df2_f = df2[df2["_key"].isin(companies_keys)].copy()

            # ── 3. Filtrar Base 3 usando claves canónicas de Base 2 ───────────
            canonical_keys = set(df2_f["_key"].dropna())
            df3 = df3.copy()
            df3["_company_key"] = df3["Compañía relacionada"].apply(norm_key)
            df3["_email"] = df3["Correo electrónico"].apply(norm_email)
            df3_f = df3[df3["_company_key"].isin(canonical_keys)].copy()

            # ── 4. Estadísticas ───────────────────────────────────────────────
            comp_b2_email = set(df2_f.loc[df2_f["_email"].notna(), "_key"])
            comp_b3_email = set(df3_f.loc[df3_f["_email"].notna(), "_company_key"])
            comp_with_email = comp_b2_email | comp_b3_email
            total_with_email = len(comp_with_email)
            total_without_email = total_companies - total_with_email

            all_emails = (
                set(df2_f.loc[df2_f["_email"].notna(), "_email"])
                | set(df3_f.loc[df3_f["_email"].notna(), "_email"])
            )
            total_unique_emails = len(all_emails)

            # ── 5. Construir Excel de salida ──────────────────────────────────
            b3_valid = df3_f[df3_f["_email"].notna()].copy()
            b3_rows = pd.DataFrame({
                "name": b3_valid["Nombre"].apply(clean_str),
                "company name": b3_valid["Compañía relacionada"].apply(clean_str),
                "email": b3_valid["_email"],
            })

            b2_valid = df2_f[df2_f["_email"].notna()].copy()
            b2_rows = pd.DataFrame({
                "name": b2_valid["Nombre mostrado"].apply(clean_str),
                "company name": b2_valid["Nombre mostrado"].apply(clean_str),
                "email": b2_valid["_email"],
            })

            # B3 primero → gana el dedup por email
            combined = pd.concat([b3_rows, b2_rows], ignore_index=True)
            combined = combined.drop_duplicates(subset=["email"], keep="first")
            combined = (
                combined.sort_values(["company name", "name"])
                .reset_index(drop=True)
            )

            # ── 6. Mostrar resultados ─────────────────────────────────────────
            st.success("✅ Proceso completado")
            st.markdown("---")
            st.markdown("### 📊 Resumen")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Empresas", total_companies)
            m2.metric("Con Correo", total_with_email)
            m3.metric("Sin Correo", total_without_email)
            m4.metric("Correos Únicos", total_unique_emails)

            found_in_b2 = set(df2_f["_key"])
            missing = companies_keys - found_in_b2
            if missing:
                with st.expander(
                    f"⚠️ {len(missing)} empresa(s) de Base 1 no encontradas en Base 2"
                ):
                    st.write(sorted(display_map.get(k, k) for k in missing))

            st.markdown("### 📋 Vista previa del Excel")
            st.dataframe(combined, use_container_width=True)

            # ── 7. Exportar Excel ─────────────────────────────────────────────
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                combined.to_excel(writer, index=False, sheet_name="Email Marketing")
            output.seek(0)

            st.download_button(
                label="⬇️ Descargar Excel",
                data=output,
                file_name="email_marketing_db.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
