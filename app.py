import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Email Marketing DB | Odoo",
    page_icon="📧",
    layout="wide"
)

st.title("📧 Generador de Base de Datos Email Marketing")
st.caption("Sube las 3 bases exportadas desde Odoo para generar tu lista de email marketing.")


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
    else:
        return pd.read_excel(file)


def norm_key(val):
    """Lowercase + strip for matching."""
    if pd.isna(val) or str(val).strip() == "":
        return None
    return str(val).strip().lower()


def norm_email(val):
    """Lowercase + strip email, returns None if empty."""
    if pd.isna(val) or str(val).strip() == "":
        return None
    return str(val).strip().lower()


def clean_str(val):
    """Strip string for display, empty string if null."""
    if pd.isna(val):
        return ""
    return str(val).strip()


def col_default(cols, keywords):
    """Return index of first column whose name contains any keyword (case-insensitive)."""
    for i, c in enumerate(cols):
        if any(k in c.lower() for k in keywords):
            return i
    return 0


# ─── File 1: Sales ────────────────────────────────────────────────────────────

st.subheader("1. Base de Ventas")
st.caption("Exportada desde el módulo de Ventas de Odoo. Debe contener al menos una columna con el nombre del cliente/empresa.")

sales_file = st.file_uploader("Subir archivo de ventas", type=["csv", "xlsx"], key="sales")
df_sales = None
sales_customer_col = None

if sales_file:
    df_sales = read_file(sales_file)
    st.dataframe(df_sales.head(5), use_container_width=True)
    cols = df_sales.columns.tolist()
    default = col_default(cols, ["cliente", "customer", "empresa", "company"])
    sales_customer_col = st.selectbox(
        "Columna de cliente / empresa:", cols, index=default, key="sc"
    )

st.divider()

# ─── File 2: Companies ────────────────────────────────────────────────────────

st.subheader("2. Base de Empresas Clientes")
st.caption("Exportada desde el módulo de Contactos de Odoo (solo empresas). Columnas esperadas: Nombre, Nombre Mostrado, Correo Electrónico.")

companies_file = st.file_uploader("Subir archivo de empresas", type=["csv", "xlsx"], key="companies")
df_companies = None
comp_name_col = None
comp_email_col = None

if companies_file:
    df_companies = read_file(companies_file)
    st.dataframe(df_companies.head(5), use_container_width=True)
    cols = df_companies.columns.tolist()
    c1, c2 = st.columns(2)
    with c1:
        comp_name_col = st.selectbox(
            "Columna de nombre empresa:",
            cols,
            index=col_default(cols, ["nombre", "name"]),
            key="cn",
        )
    with c2:
        comp_email_col = st.selectbox(
            "Columna de correo empresa:",
            cols,
            index=col_default(cols, ["correo", "email", "mail"]),
            key="ce",
        )

st.divider()

# ─── File 3: Internal Contacts ────────────────────────────────────────────────

st.subheader("3. Base de Contactos Internos")
st.caption("Exportada desde el módulo de Contactos de Odoo (personas con empresa relacionada). Columnas: Nombre, Compañía, Correo Electrónico.")

contacts_file = st.file_uploader("Subir archivo de contactos internos", type=["csv", "xlsx"], key="contacts")
df_contacts = None
cont_name_col = None
cont_company_col = None
cont_email_col = None

if contacts_file:
    df_contacts = read_file(contacts_file)
    st.dataframe(df_contacts.head(5), use_container_width=True)
    cols = df_contacts.columns.tolist()
    c1, c2, c3 = st.columns(3)
    with c1:
        cont_name_col = st.selectbox(
            "Columna de nombre contacto:",
            cols,
            index=col_default(cols, ["nombre", "name"]),
            key="tn",
        )
    with c2:
        cont_company_col = st.selectbox(
            "Columna de empresa relacionada:",
            cols,
            index=col_default(cols, ["compan", "empresa", "company"]),
            key="tc",
        )
    with c3:
        cont_email_col = st.selectbox(
            "Columna de correo contacto:",
            cols,
            index=col_default(cols, ["correo", "email", "mail"]),
            key="te",
        )

# ─── Process ──────────────────────────────────────────────────────────────────

st.divider()

all_ready = all(
    [
        df_sales is not None,
        sales_customer_col,
        df_companies is not None,
        comp_name_col,
        comp_email_col,
        df_contacts is not None,
        cont_name_col,
        cont_company_col,
        cont_email_col,
    ]
)

if not all_ready:
    st.info("Sube las 3 bases y selecciona las columnas correspondientes para habilitar el procesamiento.")
else:
    if st.button("🚀 Generar Base de Datos", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):

            # ── 1. Unique companies from sales ─────────────────────────────────
            companies_keys = set(
                k
                for k in (norm_key(c) for c in df_sales[sales_customer_col].dropna())
                if k is not None
            )
            # Map key → best display name (first occurrence preserves original casing)
            sales_name_map = {}
            for c in df_sales[sales_customer_col].dropna():
                k = norm_key(c)
                if k and k not in sales_name_map:
                    sales_name_map[k] = clean_str(c)

            total_companies = len(companies_keys)

            # ── 2. Filter companies DB ─────────────────────────────────────────
            df_comp = df_companies.copy()
            df_comp["_key"] = df_comp[comp_name_col].apply(norm_key)
            df_comp["_email"] = df_comp[comp_email_col].apply(norm_email)
            df_comp_f = df_comp[df_comp["_key"].isin(companies_keys)].copy()

            # ── 3. Filter contacts DB using Base 2 company names as reference ──
            # Base 2 and Base 3 share the same company name format (both from
            # Odoo Contacts), so we match contacts against Base 2 names, not
            # against sales names which may differ.
            canonical_keys = set(df_comp_f["_key"].dropna())

            df_cont = df_contacts.copy()
            df_cont["_company_key"] = df_cont[cont_company_col].apply(norm_key)
            df_cont["_email"] = df_cont[cont_email_col].apply(norm_email)
            df_cont_f = df_cont[df_cont["_company_key"].isin(canonical_keys)].copy()

            # ── 4. Stats ───────────────────────────────────────────────────────
            comp_with_own_email = set(
                df_comp_f.loc[df_comp_f["_email"].notna(), "_key"]
            )
            comp_with_contact_email = set(
                df_cont_f.loc[df_cont_f["_email"].notna(), "_company_key"]
            )
            comp_with_any_email = comp_with_own_email | comp_with_contact_email
            total_with_email = len(comp_with_any_email)
            total_without_email = total_companies - total_with_email

            # ── 5. Build email list ────────────────────────────────────────────
            # Contact records (priority in dedup)
            cont_email_df = df_cont_f[df_cont_f["_email"].notna()].copy()
            cont_rows = pd.DataFrame(
                {
                    "Name": cont_email_df[cont_name_col].apply(clean_str),
                    "Company Name": cont_email_df[cont_company_col].apply(clean_str),
                    "Email": cont_email_df["_email"],
                    "_source": "contact",
                }
            )

            # Company records
            comp_email_df = df_comp_f[df_comp_f["_email"].notna()].copy()
            comp_rows = pd.DataFrame(
                {
                    "Name": comp_email_df[comp_name_col].apply(clean_str),
                    "Company Name": comp_email_df[comp_name_col].apply(clean_str),
                    "Email": comp_email_df["_email"],
                    "_source": "company",
                }
            )

            # Concat: contacts first so they win the dedup
            combined = pd.concat([cont_rows, comp_rows], ignore_index=True)
            combined = combined.drop_duplicates()
            combined_dedup = (
                combined.drop_duplicates(subset=["Email"], keep="first")
                .drop(columns=["_source"])
                .sort_values(["Company Name", "Name"])
                .reset_index(drop=True)
            )

            total_emails = len(combined_dedup)

            # ── 6. Display results ─────────────────────────────────────────────
            st.success("✅ Proceso completado")
            st.markdown("---")
            st.markdown("### 📊 Resumen")

            st.markdown(
                f"""
| # | Dato | Resultado |
|---|------|-----------|
| 1 | Total de empresas (sin duplicados) | **{total_companies}** |
| 2 | Empresas **con** correo (propio o de contacto interno) | **{total_with_email}** |
| 3 | Empresas **sin** correo (ni propio ni de contacto) | **{total_without_email}** |
| 4 | Total de correos (sin duplicados) | **{total_emails}** |
"""
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Empresas", total_companies)
            m2.metric("Con Correo", total_with_email)
            m3.metric("Sin Correo", total_without_email)
            m4.metric("Total Correos", total_emails)

            # Warn about companies from sales not found in companies DB
            found_in_comp_db = set(df_comp_f["_key"])
            missing = companies_keys - found_in_comp_db
            if missing:
                with st.expander(f"⚠️ {len(missing)} empresa(s) de ventas no encontradas en la base de empresas"):
                    st.write(sorted(sales_name_map.get(k, k) for k in missing))

            st.markdown("### 📋 Vista previa del Excel")
            st.dataframe(combined_dedup, use_container_width=True)

            # ── 7. Export Excel ────────────────────────────────────────────────
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                combined_dedup.to_excel(writer, index=False, sheet_name="Email Marketing")
            output.seek(0)

            st.download_button(
                label="⬇️ Descargar Excel",
                data=output,
                file_name="email_marketing_db.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
