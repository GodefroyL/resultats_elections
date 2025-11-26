import pylatex
from pylatex import Document, Section, Subsection, Tabularx, Command, NoEscape
import lecture_resultat
import time

def generer_rapport(
        fichier='C:/Users/godef/Downloads/resultats-definitifs-par-circonscriptions-legislatives.xlsx',
        parti = ('RN', 'UXD'),
        liste_tranches = [47,45,40,35,30]):
    """
    Génère un rapport LaTeX pour le parti spécifié (ou la coalition de partis) en fonction des tranches de pourcentage données.
    un tableau est créé par deparetement avec pour chaque circonscription, les trois premières nuances avec leurs pourcentages et la tranche dans laquelle se trouve le candidat du parti étudié. Une synthèse finale donne le nombre de circonscriptions où le parti a dépassé chaque tranche.
    
    :param fichier: Fichier Excel contenant les résultats des élections législatives.
    :param parti: Nom du parti politique à analyser.
    :param liste_tranches: Liste des tranches de pourcentage à inclure dans le rapport.
    :return: Objet pylatex.Document représentant le rapport généré.
    """
    tps_total_debut = time.time() 
    doc = Document()
    doc.preamble.append(Command('title', 'Rapport des résultats par parti'))
    doc.preamble.append(Command('author', 'Analyse Électorale Automatisée par Godefroy Lecluse'))
    doc.append(Command('maketitle'))
    
    print('debut génération rapport LaTeX...')
    liste_departements = lecture_resultat.get_departements(fichier)
    n_dep = 0
    for code_dept, nom_dept in liste_departements:
        print(f'Génération du rapport pour le département {nom_dept} ({code_dept})...')
        t1 = time.time()
        with doc.create(Section(f'Département: {nom_dept} ({code_dept})')):
            with doc.create(Tabularx(f'|X|X|X|{"X|" * (len(liste_tranches) + 1)}')) as tableau:
            # Ajout ligne entete
                tableau.add_hline()
                tableau.add_row(['Nom Circonscription', 'Second tour possible', 'élu T1'] + [f' + de {liste_tranches[i]}%' for i in range(len(liste_tranches))]+ [f' - de {liste_tranches[-1]}%'])
                tableau.add_hline()
            
            # Recupération des circonscriptions du département
                liste_circos_dept = lecture_resultat.get_circonscriptions(fichier, code_dept)

                for circo_code, circo_nom in liste_circos_dept:
                # Recherche élu au premier tour
                    elu, parti_elu, score = lecture_resultat.elu_premier_tour(fichier, circo_code)
                    if elu:
                        if parti_elu in parti: x = 'X'
                        else: x = ''
                        tableau.add_row([circo_nom, f'{parti_elu} {score} élu', x] + [''] * (len(liste_tranches) + 1))
                        tableau.add_hline()
                        continue

                # Récuparation des partis présents dans la circonscription
                    partis_circo = lecture_resultat.get_partis_circonscription(fichier, circo_code)

                # Si le parti est présent, on choisit le parti de la coalition trouvé dans la liste
                    for p in parti:
                        if p in partis_circo:
                            parti = p
                            break

                # Récupération des résultats de la circonscription
                    resultats_circo = lecture_resultat.get_resultats_circo(fichier, circo_code, affichage_nom=False, affichage_parti=False)
                    second_tour_possible = []

                    for candidat_info in resultats_circo.values():
                        if float(candidat_info['pourcentage/votant'].replace('%', '').replace(',', '.')) >= 12.5:
                            second_tour_possible.append(candidat_info)
                    string_second_tour = ''
                    for candidat_info in sorted(second_tour_possible, key=lambda x: float(x['pourcentage/votant'].replace('%', '').replace(',', '.')), reverse=True): string_second_tour += f"{candidat_info.get('parti')} : {candidat_info.get('pourcentage/votant')}"

                    
                # Récupération resultats des candidats pour le parti étudié
                    pourcentage_parti_string, _, _ = lecture_resultat.get_resultats_parti(fichier, circo_code, parti)
                    pourcentage_parti = float(pourcentage_parti_string.replace('%', '').replace(',', '.'))
                    
                    tranche_atteinte = [''] * len(liste_tranches) + ['']

                    bool_tranche_atteinte = False
                    for i in range(len(liste_tranches)-1):
                        if pourcentage_parti >= liste_tranches[i]:
                            tranche_atteinte[i] = 'X'
                            bool_tranche_atteinte = True
                            break

                    if not bool_tranche_atteinte:
                        tranche_atteinte[-1] = 'X'

                    tableau.add_row([
                        circo_nom,
                        string_second_tour,
                        '',
                        ] + tranche_atteinte
                        )
                    tableau.add_hline()
            
        t2 = time.time()
        print(f'Temps de génération pour le département {nom_dept} ({code_dept}) nombre circo {len(liste_circos_dept)}: {t2 - t1:.2f} secondes')
        n_dep += 1
        print(f'Progression: {n_dep}/{len(liste_departements)} départements traités.')
    tps_total_traitement = time.time() 
    print(f'Temps total de traitement des départements: {tps_total_traitement - tps_total_debut:.2f} secondes')
    doc.generate_pdf('rapport_election', clean_tex=True)
    tps_total_fin = time.time() 
    print(f'Temps total de génération du rapport: {tps_total_fin - tps_total_debut:.2f} secondes')
    return None

if __name__ == "__main__":
    generer_rapport()
