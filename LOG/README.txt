REGISTRO UTILIZZO - RIMUOVI_PROTEZIONI_EXCEL

Il file utilizzo_tool.csv riceve una riga al termine di ogni esecuzione della macro.
Il separatore e' il punto e virgola, adatto all'importazione in Excel/Power Query.

Esiti:
- OK: tutte le protezioni di cartella e fogli sono state rimosse.
- PARZIALE: una o piu' protezioni sono rimaste attive.
- ERRORE: l'esecuzione ha incontrato un errore.

Il tracciamento non salva password, contenuti dei fogli o percorso completo del file.
Se il registro non e' raggiungibile, la macro continua comunque a funzionare.
