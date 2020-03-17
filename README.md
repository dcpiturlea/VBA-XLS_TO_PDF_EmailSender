# VBA-XLS_TO_PDF_EmailSender

Aplicatia creaza din foaia curenta un PDF pe care il salveaza autoamt intr-o locatie prestabilita(sau nu, in functie de nevoi), apoi, 
prin Outlook creaza un email prestabilit(hardcode) unde ataseaza PDF-ul salvat anterior si il trimite la ora stabilita.

Atunci cand ruleaza macro-ul principal cheama doua functii, cea care creaza pdf-ul si il saveaza si cea care trimite emailul.

In ThisWorkbook este cheama la deschidere functia privata care spune sa se astepte ora prestabilita, apoi cheama macro-ul principal
