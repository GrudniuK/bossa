
#https://www.pkotfi.pl/aktualnosci/zmiany-statutow-subfunduszy-pko-tfi/

#Dopisac do podsumowania do porównania cenê ropy
#Dopisac porownanie do poprzedniego podsumowania / tygodnia
#Dopisac aktualna cene

#https://stackoverflow.com/questions/38136008/adding-multiple-chart-series-in-quantmod-r

# https://www.analizy.pl/fundusze/fundusze-inwestycyjne/rankingi?fid=1&limit=&cid=TFIPCS&pid=PAR_PCS01&gid=&profile_send=1

if(!require(xlsx)) {install.packages("xlsx"); require(xlsx)}
if(!require(lubridate)) {install.packages("lubridate"); require(lubridate)}
if(!require(TTR)) {install.packages("TTR"); require(TTR)}
if(!require(quantmod)) {install.packages("quantmod"); require(quantmod)}
if(!require(RColorBrewer)) {install.packages("RColorBrewer"); require(RColorBrewer)}
if(!require(RCurl)) {install.packages("RCurl"); require(RCurl)}

#wczytanie bazy
setwd("G:/Wazne/Bossa")
baza <- read.xlsx("bossa.xlsx", sheetName = "baza", stringsAsFactors=FALSE)
baza$data <- as.Date(baza$data , "%Y-%m-%d")
#zaciagniecie danych z STOOQ

stooq <- vector("list", nrow(baza))

for(i in 1:nrow(baza)) {
  #i <- 1
  url_source <- paste("https://stooq.com/q/d/l/?s=",baza$ticker[i],"&i=d", sep = "")
  dane <- read.table(file = download, header = TRUE, sep = ",", na.strings = "NA", fill = TRUE, stringsAsFactors = F)
  dane$Date <- as.Date(dane$Date , "%Y-%m-%d")
  if(ncol(dane) < 6) {dane$Volume = 0}
  stooq[[i]] <- dane
  names(stooq)[i] <- baza$ticker[i]
  rm(dane, url_source); gc()
}
#save(stooq, file = "stooq.rda")
#load("stooq.rda")

stooq_por <- vector("list", nrow(baza))
for(i in 1:nrow(baza)) {
  #i <- 1
  if(!is.na(baza$porownanie[i])) {
    url_source <- paste("http://stooq.com/q/d/l/?s=",baza$porownanie[i],"&i=d", sep = "")  
    dane <- read.table(file = url_source, header = TRUE, sep = ",", na.strings = "NA", fill = TRUE, stringsAsFactors = F)
    dane$Date <- as.Date(dane$Date , "%Y-%m-%d")
    if(ncol(dane) < 6) {dane$Volume = 0}
    stooq_por[[i]] <- dane
    names(stooq_por)[i] <- baza$porownanie[i]
    rm(dane, url_source); gc()
  }
}
save(stooq_por, file = "stooq_por.rda")
#load("stooq_por.rda")

#wyniki od zakupu
for(i in 1:nrow(baza)) {
  if(exists("obecny_stan")) {obecny_stan <- rbind(obecny_stan, stooq[[i]][nrow(stooq[[i]]),])}
  if(!exists("obecny_stan")) {obecny_stan <- stooq[[i]][nrow(stooq[[i]]),]}
}

baza$data_raportu <- obecny_stan$Date
baza$wynik_pln <- obecny_stan$Close * baza$ilosc - baza$po.prowizji
baza$wynik_pr <- baza$wynik_pln / baza$po.prowizji
baza$wynik_dni <- obecny_stan$Date - baza$data.rozliczenia

#spadek od najwyzszej wartosci w okresie inwestycji
for(i in 1:nrow(baza)) {
  #i <- nrow(maksymalna_wartosc) + 1
  #rm(maksymalna_wartosc)
  walor <- stooq[[i]][stooq[[i]]$Date > baza$data.rozliczenia[i],]
  
  max <- walor[walor$Close == max(walor$Close),]
  max <- max[order(max$Date),]
  max <- max[nrow(max),]
  
  if(exists("maksymalna_wartosc")) {maksymalna_wartosc <- rbind(maksymalna_wartosc, max)}
  if(!exists("maksymalna_wartosc")) {maksymalna_wartosc <- max}
  rm(walor, max); gc()
}

baza$data_max <- maksymalna_wartosc$Date
baza$od_max_pln <- ( obecny_stan$Close - maksymalna_wartosc$Close ) * baza$ilosc
baza$od_max_pr <- ( obecny_stan$Close - maksymalna_wartosc$Close ) / maksymalna_wartosc$Close
baza$od_max_dni <- ( obecny_stan$Date - maksymalna_wartosc$Date )

#wyniki tygodniowe

for(i in 1:nrow(baza)) {
#i <- 1
  #walor <- stooq[[i]][stooq[[i]]$Date > baza$data.rozliczenia[i],]
  walor <- stooq[[i]]
  walor$weekday <- weekdays(walor$Date, abbreviate = T)
  walor <- unique( rbind( walor[walor$weekday == "pt.",], walor[nrow(walor),]) )
  walor$prev <- c(NA, walor[1:(nrow(walor)-1),"Close"])
  walor$pr <- ( walor$Close - walor$prev ) / walor$Close
  walor$week <- paste0("week_", nrow(walor) - row(walor)[,1] + 1 )
  walor <- tail(walor, 5)
  
  wyniki_tygodniowe_add <- as.data.frame(t(walor$pr))
  colnames(wyniki_tygodniowe_add) <- walor$week

  if(exists("wyniki_tygodniowe")) {wyniki_tygodniowe <- rbind(wyniki_tygodniowe, wyniki_tygodniowe_add)}
  if(!exists("wyniki_tygodniowe")) {wyniki_tygodniowe <- wyniki_tygodniowe_add}
  rm(walor, wyniki_tygodniowe_add); gc()
}

baza <- cbind(baza, wyniki_tygodniowe)

#zapisz podsumowanie 
date_file <- year(max(obecny_stan$Date)) * 10000 + month(max(obecny_stan$Date)) * 100 + day(max(obecny_stan$Date))
file_name <- paste0("raport_", date_file, ".xlsx", collapse = "")
write.xlsx(baza, file_name, sheetName = "all", row.names = F, append = T)

#podsumowania na dla tickerow - na oddzielnych zakladkach
for(i in 1:nrow(baza)) {write.xlsx(baza[i,], file_name, sheetName = baza$ticker[i], col.names = T, row.names = F, append = T)}

#wykresy
for(i in 1:nrow(baza)) {

data <- stooq[[i]][stooq[[i]]$Date >= baza$data.rozliczenia[i] - 150,]
data_xts <- xts(x=data[,-1], order.by = data$Date)

colShort <- brewer.pal(n = 9, name = "Purples")[4:9]
colLong <- brewer.pal(n = 9, name = "Reds")[4:9]

addGuppy <- newTA(
  on=-1,
  FUN=GMMA,
  preFUN=Cl,
  col=c(colShort,colLong),
  legend='GMMA'
)

dc <- DonchianChannel(cbind(Hi(data_xts), Lo(data_xts)), n = 30)

png("wykres1.png", width = 1000, height = 500)
p <- chartSeries(
  data_xts,
  major.ticks = "months", 
  TA='
  addTA(dc$low, on=1, col="red", lty=1, lwd=2);
  addTA(dc$high, on=1, col="green", lty=1, lwd=2);
  addVo();
  addLines(h=baza$cena[i] * 1.1, on=-1, col = "green");
  addLines(h=baza$cena[i], on=-1, col = "blue");
  addLines(h=baza$cena[i] * 0.95, on=-1, col = "red");
  addLines(v=which(data$Date == baza$data.rozliczenia[i]), on=-1, col = "blue");
  addLines(h=obecny_stan$Close[i], on=-1, col = "yellow");
  '
)
print(p)
dev.off()

png("wykres2.png", width = 1000, height = 500)
p <- chartSeries(
  data_xts,
  major.ticks = "months", 
  TA='
  addLines(h=baza$cena[i] * 1.1, on=-1, col = "green");
  addLines(h=baza$cena[i], on=-1, col = "blue");
  addLines(h=baza$cena[i] * 0.95, on=-1, col = "red");
  addLines(v=which(data$Date == baza$data.rozliczenia[i]), on=-1, col = "blue");
  addLines(h=obecny_stan$Close[i], on=-1, col = "yellow");
  addGuppy();
  addMACD();
  addRSI();
  '
)
print(p)
dev.off()

#porownanie
if(!is.na(baza$porownanie[i])) {
  data <- stooq[[i]][stooq[[i]]$Date >= baza$data.rozliczenia[i],]
  data_por <- stooq_por[[i]][stooq_por[[i]]$Date >= baza$data.rozliczenia[i],]

  #normalizacja
  data$Close <- data$Close / data$Close[151]
  data_por$Close <- data_por$Close / data_por$Close[151]
  
  data_xts <- xts(x=data[,-1], order.by = data$Date)
  data_xts_por <- xts(x=data_por[,-1], order.by = data_por$Date)

  png("wykres3.png", width = 1000, height = 500)
  p <- chartSeries(Cl(data_xts), TA='
    addTA(Cl(data_xts_por), on=1, col = "red");
    addLines(v=which(data$Date == baza$data.rozliczenia[i]), on=-1, col = "blue");
  '
  )
  print(p)
  dev.off()
}

wb <- loadWorkbook(file_name)
sheets <- getSheets(wb)
sheet <- sheets[[i+1]] #bo pierwsza to podsumowanie
addPicture(file = "wykres1.png", sheet = sheet, startColumn = 1, startRow = 5)
addPicture(file = "wykres2.png", sheet = sheet, startColumn = 1, startRow = 30)
if(!is.na(baza$porownanie[i])) {
  addPicture(file = "wykres3.png", sheet = sheet, startColumn = 1, startRow = 55)
}
saveWorkbook(wb, file_name)
unlink("wykres1.png")
unlink("wykres2.png")
unlink("wykres3.png")
rm(data, data_xts, dc, p, wb); gc()
}

date_file
traceback()

# https://www.analizy.pl/fundusze/fundusze-inwestycyjne/notowania/firma/TFIPCS/produkt/PAR_PCS01/grupa/--/nazwa/--/typ-jednostki/podstawowe/typ/last/data/--/limit/100/strona/1/sort/1/sort_dir/ASC/
