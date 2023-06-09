import locale
import zipfile
from io import BytesIO

import markdown
import pandas as pd
import streamlit as st
from weasyprint import HTML

locale.setlocale(locale.LC_ALL, locale="fr_FR")

def surveillances_enseignant_pdf(data, enseignant, modèle_convocation, style_sheet="style_convocation.css"):
	surveillances_enseignant = data.loc[
	    data["Enseignant"] == enseignant, ["Date", "Horaire", "Matière"]
	].sort_values(by=["Date", "Horaire"])
	surveillances_enseignant["Date"] = surveillances_enseignant[
	    "Date"
	].dt.strftime("%A %d %B %Y")

	rendered_md = modèle_convocation.format(
		enseignant=enseignant,
		surveillances=surveillances_enseignant.to_markdown(
		    index=False
		)
	)
	rendered_html = markdown.markdown(rendered_md, extensions=["tables"])
	pdf = HTML(string=rendered_html).write_pdf(
	    stylesheets=[style_sheet]
	)
	return pdf

def surveillants_epreuve_pdf(data, epreuve, modèle_fiche, style_sheet="style_fiche.css"):
    surveillants = data.loc[
        data["VraiMatière"] == epreuve, ["Enseignant"]
    ].sort_values(by=["Enseignant"])
    surveillants["Salle"] = ""
    surveillants["Observation"] = ""

    rendered_md = modèle_fiche.format(
            epreuve=epreuve,
            surveillants=surveillants.to_markdown(index=False),
            date=data.loc[data["VraiMatière"] == epreuve, ["Date"]]
            .iloc[0, 0]
            .strftime("%A %-d %B %Y"),
            horaire=data.loc[
                data["VraiMatière"] == epreuve, ["Horaire"]
            ].iloc[0, 0]
    )
    rendered_html = markdown.markdown(rendered_md, extensions=["tables"])
    pdf = HTML(string=rendered_html).write_pdf(
        stylesheets=[style_sheet]
    )
    return pdf

st.title("Gestion des surveillances")

st.write(
    """
    Cette application a été conçue dans le but de faciliter la
    génération des documents nécessaires à la gestion des surveillants des contrôles et examens
"""
)

st.subheader("1. 📂 Uploader un fichier .ods")

if uploaded := st.file_uploader(
    "Fichier .ods contenant les données brut des surveillances", type="ods"
):
    data = pd.read_excel(uploaded)
    st.header("1. Générateur de convocations")
    st.subheader("1.1 📑 Créer un modèle de convocation")
    st.markdown(
        """Utiliser ou modifier le modèle de convocation ci-dessous. Utiliser les balises :
- *{enseignant}* sera remplacée par le **nom de l'enseignant**
- *{surveillances}* sera remplacée par le **tableau de surveillances**
"""
    )
    modèle_convocation = st.text_area(
        "Modèle de la convocation",
        value="""# Convocation

Mme./Mlle./Mr. **{enseignant}**

Vous êtes cordialement invité(e) à assurer les surveillances
des **examens semestriels** selon le planning ci-dessous :

{surveillances}

**Le chef de département CPST**

<span style="color:red; font-size:.8em">Présence obligatoire
dans le hall entre les deux amphis 10 minutes avant le début de chaque épreuve</span>""",
        height=330,
        help="Doit respecter la syntaxe de Markdown",
    )
    with st.expander("Aperçu de la convocation"):
        st.markdown(modèle_convocation, unsafe_allow_html=True)
    st.subheader("1.2. ⚙️ Générer les convocations")
    st.markdown(
        "Cliquer sur le bouton ```⚡️ Générer les convocation ...``` et attendre que le processus se termine."
    )
    if st.button("⚡️ Générer les convocation ..."):
        enseignants = sorted(data["Enseignant"].unique())
        progress_bar = st.progress(0)
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_STORED) as zip_file:
            for index, enseignant in enumerate(enseignants):
                pdf = surveillances_enseignant_pdf(data, enseignant, modèle_convocation)
                zip_file.writestr(f"{enseignant}.pdf", pdf)
                progress_bar.progress(
                    (index + 1) / len(enseignants),
                    f"[{index+1}/{len(enseignants)}] Convocation de {enseignant}",
                )
        st.download_button(
            label="⬇️ Télécharger les convocations",
            data=zip_buffer,
            file_name="convocations.zip",
            mime="application/zip",
        )
    st.header("2. Générateur de fiches de suivi")
    st.subheader("2.1. 📑 Créer un modèle de fiche de suivi")
    st.markdown(
        """Utiliser ou modifier le modèle de fiche de suivi ci-dessous. Utiliser les balises :
- *{date}* sera remplacée par la **date de l'épreuve**
- *{épreuve}* sera remplacée par le **nom de l'épreuve**
- *{horaire}* sera remplacée par l'**horaire de l'épreuve**
- *{surveillants}* sera remplacée par le tableau des **surveillants**
"""
    )
    modèle_fiche = st.text_area(
        "Modèle de fiche de suivi",
        value="""## Suivi des surveillants

- **Date :** *{date}*
- **Matière :** *{epreuve}*
- **Horaire :** *{horaire}*

{surveillants}""",
        height=190,
        help="Doit respecter la syntaxe de Markdown",
    )
    with st.expander("Aperçu de la fiche de suivi"):
        st.markdown(modèle_fiche, unsafe_allow_html=True)

    st.subheader("2.2. ⚙️ Générer les fiches de suivi")
    st.markdown(
        "Cliquer sur le bouton ```⚡️ Générer les fiches de suivi ...``` et attendre que le processus se termine."
    )
    if st.button("⚡️ Générer les fiches de suivi ..."):
        epreuves = sorted(data["VraiMatière"].unique())
        progress_bar = st.progress(0)
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_STORED) as zip_file:
            for index, epreuve in enumerate(epreuves):
                pdf = surveillants_epreuve_pdf(data, epreuve, modèle_fiche)
                zip_file.writestr(f"{epreuve}.pdf".replace("/", "-"), pdf)
                progress_bar.progress(
                    (index + 1) / len(epreuves),
                    f"[{index+1}/{len(epreuves)}] Fiche de suivi de {epreuve}",
                )
        st.download_button(
            label="⬇️ Télécharger les fiches de suivis",
            data=zip_buffer,
            file_name="fiches_de_suivi.zip",
            mime="application/zip",
        )
