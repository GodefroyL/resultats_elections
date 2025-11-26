from pylatex import Document, Section, LongTabularx, Command, NoEscape, Package
import lecture_resultat
import time

def generer_rapport(
        fichier='C:/Users/godef/Downloads/resultats-definitifs-par-circonscriptions-legislatives.xlsx',
        parti_etudie=('RN', 'UXD'),
        liste_tranches=[47, 45, 43, 40, 35, 30]):
    tps_total_debut = time.time()
    doc = Document()

    # Ajout des packages et des commandes de mise en page
    doc.preamble.append(Package('ltablex'))
    doc.preamble.append(Package('array'))
    doc.preamble.append(Package('babel', options=['french']))
    doc.preamble.append(Package('geometry', options=['left=1cm', 'right=1cm', 'top=2cm', 'bottom=2cm']))

    doc.preamble.append(Command('title', f'Rapport des résultats électoraux pour le parti {" / ".join(parti_etudie)}'))
    doc.preamble.append(Command('author', 'Analyse Électorale Automatisée par Godefroy Lecluse'))
    doc.append(Command('maketitle'))

    resultats_election = lecture_resultat.lecture_resultat(fichier)
    liste_departements = lecture_resultat.get_departements(resultats_election)

    for code_dept, nom_dept in liste_departements:
        with doc.create(Section(f'Département: {nom_dept} ({code_dept})')):
            taille_cellule_tranche = str(10/(len(liste_tranches)+1)) + 'cm'
            
            with doc.create(LongTabularx(f'|p{{1.5cm}}|p{{2.3cm}}|X|{f"p{{{taille_cellule_tranche}}}|" * (len(liste_tranches) + 1)}', width_argument=NoEscape(r'\textwidth'))) as tableau:
            
            # Ajout ligne en-tête
                tableau.add_hline()
                tableau.add_row(['Circo', '2nd tour', 'élu T1'] + [f' + de {liste_tranches[i]}%' for i in range(len(liste_tranches))] + [f' - de {liste_tranches[-1]}%'])
                tableau.add_hline()

            # Récupération des circonscriptions du département
                liste_circos_dept = lecture_resultat.get_circonscriptions(resultats_election, code_dept)
                for circo_code, circo_nom in liste_circos_dept:
                    circo_nom = circo_nom.replace('nscription', '')

                # Recherche élu au premier tour
                    elu, parti_elu, score = lecture_resultat.elu_premier_tour(resultats_election, circo_code)
                    if elu:
                        if parti_elu in parti_etudie:
                            x = 'X'
                        else:
                            x = ''
                        tableau.add_row([circo_nom, f'{parti_elu} : {score} élu', x] + [''] * (len(liste_tranches) + 1))
                        tableau.add_hline()

                    else:
                    # Pas d'élu au premier tour
                    # Récupération des partis présents dans la circonscription
                        partis_circo = lecture_resultat.get_partis_circonscription(resultats_election, circo_code)
                        # Si le parti est présent, on choisit le parti de la coalition trouvé dans la liste
                        for p in parti_etudie:
                            if p in partis_circo:
                                parti = p
                            else:
                                parti = parti_etudie[0]

                    # Récupération des résultats de la circonscription
                        resultats_circo = lecture_resultat.get_resultats_circo(resultats_election, circo_code, affichage_nom=False, affichage_parti=False)

                    # Recherche des candidats pouvant aller au second tour
                        second_tour_possible = []
                        for candidat_info in resultats_circo.values():
                            if float(candidat_info['pourcentage/votant'].replace('%', '').replace(',', '.')) >= 12.5:
                                second_tour_possible.append(candidat_info)
                        string_second_tour = ''
                        for candidat_info in sorted(second_tour_possible, key=lambda x: float(x['pourcentage/votant'].replace('%', '').replace(',', '.')), reverse=True):
                            string_second_tour += f"{candidat_info.get('parti')} : {candidat_info.get('pourcentage/votant')} \n"

                    # Récupération des résultats des candidats pour le parti étudié
                        pourcentage_parti_string, _, _ = lecture_resultat.get_resultats_parti(resultats_election, circo_code, parti)
                        pourcentage_parti = float(pourcentage_parti_string.replace('%', '').replace(',', '.'))

                    # Remplissage du tableau avec les tranches atteintes
                        tranche_atteinte = [''] * len(liste_tranches) + ['']
                        bool_tranche_atteinte = False
                        for i in range(len(liste_tranches)):
                            if pourcentage_parti >= liste_tranches[i]:
                                tranche_atteinte[i] = 'X'
                                bool_tranche_atteinte = True
                                break
                        if not bool_tranche_atteinte:
                            tranche_atteinte[-1] = 'X'

                    # Ajout de la ligne au tableau
                        tableau.add_row([circo_nom, string_second_tour, ''] + tranche_atteinte)
                        tableau.add_hline()

    tps_total_traitement = time.time()
    print(f'Temps total de traitement des départements: {tps_total_traitement - tps_total_debut:.2f} secondes')

# Génération du PDF
    doc.generate_pdf(f'rapport_election_{"-".join(parti_etudie)}', clean_tex=True)

    tps_total_fin = time.time()
    print(f'Temps total de génération du rapport: {tps_total_fin - tps_total_debut:.2f} secondes')

if __name__ == '__main__':
    generer_rapport()
