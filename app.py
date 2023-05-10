import streamlit as st
import pandas as pd
from jinja2 import Environment, BaseLoader
import markdown
import zipfile
from io import BytesIO
import locale

locale.setlocale(locale.LC_ALL, locale="fr_FR")

st.title("Gestion des surveillances")

st.write(
    """
    Cette application a été conçu dans le but de faciliter la
    génération des documents nécessaires à la gestion des surveillants des contrôles et examens
"""
)

st.header("Générateur de convocations")
st.subheader("1. 📂 Uploader un fichier .ods")

if uploaded := st.file_uploader(
    "Fichier .ods contenant les données brut des surveillances", type="ods"
):
    data = pd.read_excel(uploaded)
    st.subheader("2. 📑 Créer un modèle de convocation")
    st.markdown("""Utiliser ou modifier le modèle de convocation ci-dessous. Utiliser les balises :
- *{{ enseignant }}* sera remplacée par le **nom de l'enseignant**
- *{{ surveillances }}* sera remplacée par le **tableau de surveillances**
"""
    )
    modèle_convocation = st.text_area(
        "Modèle de la convocation",
        value="""# Convocation

Mr./Mme./Mlle. **{{ enseignant }}**

Vous êtes cordialement invité(e) à assurer les surveillances
des **examens semestriels** selon le planning ci-dessous :

{{ surveillances }}

**Le chef de département CPST**

<span style="color:red; font-size:.8em">Présence obligatoire
dans le hall entre les deux amphis 10 minutes avant le début de chaque épreuve</span>""",
        height=330,
        help="Doit respecter la syntaxe de Markdown",
    )
    with st.expander("Aperçu de la convocation"):
        st.markdown(modèle_convocation, unsafe_allow_html=True)
    st.subheader("3. ⚙️ Générer les convocations")
    st.markdown(
        "Cliquer sur le bouton ```⚡️ Générer les convocation ...``` et attendre que le processus se termine."
    )
    if st.button("⚡️ Générer les convocation ..."):
        enseignants = sorted(data["Enseignant"].unique())
        progress_bar = st.progress(0)
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_LZMA) as zip_file:
            for index, enseignant in enumerate(enseignants):
                surveillances_enseignant = data.loc[
                    data["Enseignant"] == enseignant, ["Date", "Horaire", "Matière"]
                ].sort_values(by=["Date", "Horaire"])
                surveillances_enseignant["Date"] = surveillances_enseignant[
                    "Date"
                ].dt.strftime("%A %d %B %Y")
                convocation_template = Environment(loader=BaseLoader).from_string(
                    modèle_convocation
                )
                rendered_md = convocation_template.render(
                    {
                        "enseignant": enseignant,
                        "surveillances": surveillances_enseignant.to_markdown(
                            index=False
                        ),
                    }
                )
                rendered_html = markdown.markdown(rendered_md, extensions=["tables"])
                pdf = HTML(string=rendered_html).write_pdf(
                    stylesheets=["style_convocation.css"]
                )
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

    st.subheader("4. 📑 Créer un modèle de fiche de suivi")
    st.markdown("""Utiliser ou modifier le modèle de fiche de suivi ci-dessous. Utiliser les balises :
- *{{ date }}* sera remplacée par la **date de l'épreuve**
- *{{ épreuve }}* sera remplacée par le **nom de l'épreuve**
- *{{ horaire }}* sera remplacée par l'**horaire de l'épreuve**
- *{{ surveillants }}* sera remplacée par le tableau des **surveillants**
"""
    )
    modèle_fiche = st.text_area(
        "Modèle de fiche de suivi",
        value="""## Suivi des surveillants

- **Date :** *{{ date }}*
- **Matière :** *{{ epreuve }}*
- **Horaire :** *{{ horaire }}*

{{ surveillants }}""",
        height=190,
        help="Doit respecter la syntaxe de Markdown",
    )
    with st.expander("Aperçu de la fiche de suivi"):
        st.markdown(modèle_fiche, unsafe_allow_html=True)
    
    st.subheader("5. ⚙️ Générer les fiches de suivi")
    st.markdown(
        "Cliquer sur le bouton ```⚡️ Générer les fiches de suivi ...``` et attendre que le processus se termine."
    )
    if st.button("⚡️ Générer les fiches de suivi ..."):
        epreuves = sorted(data["VraiMatière"].unique())
        progress_bar = st.progress(0)
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_LZMA) as zip_file:
            for index, epreuve in enumerate(epreuves):
                surveillants = data.loc[
                        data["VraiMatière"] == epreuve, ["Enseignant"]
                    ].sort_values(by=["Enseignant"])
                surveillants['Salle'] = ''
                surveillants['Observation'] = ''
                fiche_template = Environment(loader=BaseLoader).from_string(
                    modèle_fiche
                )
                rendered_md = fiche_template.render(
                    {
                        "epreuve": epreuve,
                        "surveillants": surveillants.to_markdown(index=False),
                        "date" : data.loc[data['VraiMatière'] == epreuve, ['Date']].iloc[0,0].strftime("%A %-d %B %Y"),
                        "horaire" : data.loc[data['VraiMatière'] == epreuve, ['Horaire']].iloc[0,0]
                    }
                )
                rendered_html = markdown.markdown(rendered_md, extensions=["tables"])
                pdf = HTML(string=rendered_html).write_pdf(
                    stylesheets=["style_fiche.css"]
                )
                zip_file.writestr(f"{epreuve}.pdf".replace("/","-"), pdf)
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