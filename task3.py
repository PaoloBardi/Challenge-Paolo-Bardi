import xlsxwriter
import csv

with open("./data.csv", newline="") as fileDati:
    csvHandler = csv.reader(fileDati,delimiter=",")
    listaDocenti = []
    corsiTrovati = []
    corsoAttuale = ""
    rigaXls=0
    colonnaXls=0
    
    numeroDocenti = int(input("Inserisci il numero di docenti da ricercare [0 to view all]: "))
    
    while len(listaDocenti) < numeroDocenti:
        docenteAttuale = input("Inserisci docente [Name e/o Surname]: ")
        listaDocenti.append(docenteAttuale)
        
    filtroAnnoSolare = input("Desideri filtrare per anno Solare? [Y/N]: ")
    if filtroAnnoSolare=="Y":
        annoSolare = input("Inserisci anno solare: ")
    else:
        pass
    
    filtroAnnoCorso = input("Desideri filtrare per anno di corso? [Y/N]: ")
    if filtroAnnoCorso=="Y":
        annoCorso = input("Inserisci anno di corso da ricercare: ")
    else:
        pass

    salvataggioXls = input("Desideri salvare il risultato della ricerca in un file Xls? [Y/N]: ")
    if salvataggioXls=="Y":
        documentoUscita=xlsxwriter.Workbook("risultati ricerca.xlsx")
        paginaDocumento=documentoUscita.add_worksheet()
        paginaDocumento.write(rigaXls,colonnaXls,"Numero Corso")
        paginaDocumento.write(rigaXls,colonnaXls+1,"Nome Corso")
        paginaDocumento.write(rigaXls,colonnaXls+2,"Istituto")
        rigaXls+=1

    if numeroDocenti==0:
        next(csvHandler)
        if filtroAnnoSolare=="N" and filtroAnnoCorso=="N":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler]     
            
    
        if filtroAnnoSolare=="N" and filtroAnnoCorso=="Y":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if annoCorso==riga[6]]
            

        if filtroAnnoSolare=="Y" and filtroAnnoCorso=="N":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if annoSolare in riga[2]]
            

        if filtroAnnoSolare=="Y" and filtroAnnoCorso=="Y":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if annoSolare in riga[2] and annoCorso==riga[6]]

        for risultato in dati:
            corsoAttuale=str(risultato[0])
            if corsoAttuale not in corsiTrovati:
                print(f"Numero corso: {risultato[0]} -- Nome corso: {risultato[1]} -- Istituto: {risultato[2]}")
                if salvataggioXls=="Y":
                    paginaDocumento.write(rigaXls,colonnaXls,risultato[0])
                    paginaDocumento.write(rigaXls,colonnaXls+1,risultato[1])
                    paginaDocumento.write(rigaXls,colonnaXls+2,risultato[2])
                    rigaXls+=1
                corsiTrovati.append(corsoAttuale)
        fileDati.seek(0)


    for docente in listaDocenti:
        if filtroAnnoSolare=="N" and filtroAnnoCorso=="N":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if docente in riga[4]]            
            
    
        if filtroAnnoSolare=="N" and filtroAnnoCorso=="Y":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if docente in riga[4] and annoCorso==riga[6]]
            

        if filtroAnnoSolare=="Y" and filtroAnnoCorso=="N":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if docente in riga[4] and annoSolare in riga[2]]
            

        if filtroAnnoSolare=="Y" and filtroAnnoCorso=="Y":
            dati = [(riga[1],riga[3],riga[0]) for riga in csvHandler if docente in riga[4] and annoSolare in riga[2] and annoCorso==riga[6]]

        for risultato in dati:
            corsoAttuale=str(risultato[0])
            if corsoAttuale not in corsiTrovati:
                print(f"Numero corso: {risultato[0]} -- Nome corso: {risultato[1]} -- Istituto: {risultato[2]}")
                if salvataggioXls=="Y":
                    paginaDocumento.write(rigaXls,colonnaXls,risultato[0])
                    paginaDocumento.write(rigaXls,colonnaXls+1,risultato[1])
                    paginaDocumento.write(rigaXls,colonnaXls+2,risultato[2])
                    rigaXls+=1            
                corsiTrovati.append(corsoAttuale)
        fileDati.seek(0)

    if salvataggioXls=="Y":
        documentoUscita.close()
  

        
        
            
                    
            
            
                
            

    
            


