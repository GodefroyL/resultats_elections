# from openpyxl import load_workbook

# # Charger le fichier Excel
# fichier = 'C:/Users/godef/Downloads/resultats-definitifs-par-circonscriptions-legislatives.xlsx'
# wb = load_workbook(filename=fichier, data_only=True)

# # Sélectionner la feuille active (ou spécifier le nom de la feuille)
# feuille = wb.active

# circonscriptions = {1,2}

# infos = {}


# # Parcourir chaque ligne
# for ligne in feuille.iter_rows(values_only=True):
#     # print(ligne)  # Affiche chaque ligne sous forme de tuple
#     a = input("Appuyez sur Entrée pour continuer...")  # Pause après chaque ligne
#     if a.lower() == 'q':  # Permet de quitter la boucle si l'utilisateur entre 'q'
#         break
#     i=0
#     for element in ligne:
#         print(f'ligne[{i}] : {element}')  # Affiche chaque élément de la ligne
#         i+=1
#         b = input("")  # Pause après chaque élément
#         if b.lower() == 'q':  # Permet de quitter la boucle si l'utilisateur entre 'q'
#             break

string = 'Bonjour\nComment ça va ?\n'
print(string)