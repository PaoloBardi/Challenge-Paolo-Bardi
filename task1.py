import csv

with open("./data.csv", newline="") as fileDati:
    csvHandler = csv.reader(fileDati,delimiter=",")
    docenteInserito = input("Inserisci il docente [Name Surname]: ")
    if docenteInserito=="":
            next(csvHandler)
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler]
    else:
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if docenteInserito in riga[4]]       
    for risultato in dati:
            print(f"Numero corso: {risultato[0]} -- Nome corso: {risultato[1]} -- Istituto: {risultato[2]}")
            
            



        
