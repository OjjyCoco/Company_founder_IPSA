import pandas as pd

infini = 1
print("----")
print("AUTEUR : Nathan Heckmann a.k.a Hackermann, IPSA Paris Promo 2024")
print("Présentation du programme:")
print("Ce programme utilise 4 fichiers excels contenant les entreprises dont les anciennes promo de l'IPSA ont obtenues un stage. Ce programme vous permettra de trier ces entreprises en fonction de leur localisation. Ce tri s'effectue par l'entrée d'un ou plusieurs départements.")
print("Ci-dessous, entrez d'abord le nombre de départements dans lesquels vous voulez trouver les sociétés, puis appuyer sur la touche Entrée, avant de saisir chaque département.")
print("----")
while infini == 1:
    mes_cp = []
    pays = input("Pays de la recherche: ").upper()
    nbr_dep = int(input("Nombre de départements à filtrer: "))
    for j in range(nbr_dep):
        d = str(input("Deux premiers chiffres du code postale du département " + str(j+1) + " (ex: 75): "))
        mes_cp.append(d)


    def find_entreprise_dans_dep(dep, file):
        df = pd.read_excel(file)

        i = 0
        liste_index = []
        for e in df.SOCIETE_SIGNATAIRE_CP:
            if df["SOCIETE_SIGNATAIRE_PAYS"][i] == pays:
                if str(e) != "nan":
                    if str(e)[:2] in dep:
                        liste_index.append(i)
            i += 1

        for index in liste_index:   
            print("------------------")
            print("Société: ",df.SOCIETE_SIGNATAIRE_LIB.tolist()[index])
            print("Ville: ", df.SOCIETE_SIGNATAIRE_VILLE.tolist()[index])
            print("Département: ", df.SOCIETE_SIGNATAIRE_CP.tolist()[index])
            print("Secteur d'activité (nan = pas mentionné): ", df.SECTEUR_ACTIVITE.tolist()[index])


    for n in range(5, 10):
        print("")
        print("*****************************************************")
        print("Extraction stage 201"+str(n)+": ")
        find_entreprise_dans_dep(mes_cp, "extraction_stages_201"+str(n)+".xlsx")
    print("-")
    print("-")
    infini = int(input("Entrez 1 pour relancer une recherche et 0 pour arrêter le programme: "))
