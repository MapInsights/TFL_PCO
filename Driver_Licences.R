#Weekly Drivers and vehicles

#setwd('S:/Data Analytics/Analysis/Drivers/TFL_PCO')
setwd('C:/Programs/gtc_tasks/TfL_PCO')
wd<-getwd()

date <- as.Date(Sys.time())

# load package for sql
library(DBI)
library(RODBC)
# connect to database
odbcChannel <- odbcConnect('Rstudio', uid='Daria Alekseeva', pwd='Welcome30')

Lite<-sqlQuery( odbcChannel, "--Only DD drivers, leave the null for GTC and this has been combined with vehicles.
                select distinct(mmph.actorId), --d.employee_id,
                i.firstName,i.surname,--c.name,d.vehicle_id ,d.insurance_no,
                -- pco.id,
                dcf.number 'PCO Number',--dcf.type
                v.reg_number,
                makes.name 'make',
                models.name 'model'
                --pco.expiration_date,
                --pco.type,
                --pco.documentType,
                --anto.maxdate
                from echo_core_prod..mc_mobile_presence_history mmph
                left join echo_core_prod..drivers d on d.mobileId=mmph.actorId
                --	left join (select pco2.driver_id, max(pco2.creation_date) 'maxdate'  
                --		from driver_pco_licenses pco2
                --		where pco2.type=0
                --		group by pco2.driver_id )anto
                --		on anto.driver_id = d.employee_id
                --	left join driver_pco_licenses pco on pco.driver_id=d.employee_id and pco.type=0 and pco.documentType=1 and pco.creation_date=anto.maxdate
                left join echo_core_prod..document_custom_fields dcf on dcf.driver_id=d.employee_id and dcf.type like '%PCO%'  and (dcf.deleted = 0 or dcf.deleted is null)
                left join echo_core_prod..callsigns c on c.driver_id = d.employee_id
                left join echo_core_prod..individuals i on i.id = d.employee_id 
                left join echo_core_prod..vehicles v on v.id = d.vehicle_id
                left join echo_core_prod..makes on makes.id= v.make_id
                left join echo_core_prod..models on models.id=v.model_id
                where mmph.actorPresence = 'status.online'
                -- and timestamp between '2016-12-26' and '2017-01-01 23:59:59'
                and datepart(ISO_WEEK,mmph.timestamp)=   datepart(ISO_WEEK,getdate()) - 1 
                and c.name like '%DD%'
                "
                
)


GTC_Drivers<- sqlQuery(odbcChannel, "                      
                       select distinct(mmph.actorId),
                       i.firstName,i.surname,c.name,
                       dcf.number 'DCF_PCO_Number',
                       v.reg_number ,m.name as 'MAKE'
                       from echo_core_prod..mc_mobile_presence_history mmph
                       left join echo_core_prod..drivers d on d.mobileId=mmph.actorId
                       left join (select pco2.driver_id, max(pco2.creation_date) 'maxdate'  
                       from echo_core_prod..driver_pco_licenses pco2
                       where pco2.type=0
                       group by pco2.driver_id )anto
                       on anto.driver_id = d.employee_id
                       left join echo_core_prod..driver_pco_licenses pco on pco.driver_id=d.employee_id and pco.type=0 and pco.documentType=1 and pco.creation_date=anto.maxdate
                       left join echo_core_prod..callsigns c on c.driver_id = d.employee_id
                       left join echo_core_prod..vehicles v on v.id = d.vehicle_id
                       left join echo_core_prod..makes m on m.id = v.make_id
                       left join echo_core_prod..individuals i on i.id = d.employee_id
                       left join echo_core_prod..document_custom_fields dcf on dcf.driver_id=d.employee_id and dcf.type like '%PCO%'  and (dcf.deleted = 0 or dcf.deleted is null)
                       where mmph.actorPresence = 'status.online'
                       and datepart(ISO_WEEK,mmph.timestamp)=   datepart(ISO_WEEK,getdate()) - 1 
                       and c.name not like '%DD%' and actorId != '111111'"
                       
)

odbcClose(odbcChannel)

#clean up
#remove Test Driver
GTC_Drivers2<-GTC_Drivers[!grepl('GTC',GTC_Drivers$firstName),]
Lite2<-Lite[!grepl('GTC',Lite$firstName),]


GTC_Drivers2<-GTC_Drivers[!grepl('Test',GTC_Drivers$firstName),]
Lite2<-Lite[!grepl('Test',Lite$firstName),]


#remove last 4 digits

substrRight <- function(x, n){
  substr(x, nchar(x)-n+1, nchar(x))
}
substrLeft <- function(x, n){
  substr(x, 0,nchar(x)-n)
}

#GTC first
GTC_Drivers2$left<-substrLeft(as.character(GTC_Drivers2$'DCF_PCO_Number'),4)
GTC_Drivers2$right<-substrRight(as.character(GTC_Drivers2$'DCF_PCO_Number'),4)




GTC_Drivers2$PCO2<-gsub('0.0.','',GTC_Drivers2$right)

GTC_Drivers2$PCO_Number<-ifelse(nchar(GTC_Drivers2$left)<3,
                                paste(GTC_Drivers2$left,GTC_Drivers2$right,sep=''),
                                paste(GTC_Drivers2$left,GTC_Drivers2$PCO2,sep=''))




GTC_Final<-GTC_Drivers2[,c(1:7)]

#GTC_Final[GTC_Final$PCO_Number=="NANA",]$PCO_Number<-""
#GTC_Final[GTC_Final$PCO_Number=="NANA",]$PCO_Number<-""


#Now Lite
Lite2$left<-substrLeft(as.character(Lite2$`PCO Number`),4)
Lite2$right<-substrRight(as.character(Lite2$`PCO Number`),4)



Lite2$PCO2<-gsub('0.0.','',Lite2$right)


Lite2$PCO_Number<-ifelse(nchar(Lite2$left)<3,
                         paste(Lite2$left,Lite2$right,sep=''),
                         paste(Lite2$left,Lite2$PCO2,sep=''))

Lite_Final<-Lite2[,c(1:7)]
#Lite_Final<-Lite2[,c(1:3,9)]



library(xlsx)
fullfile<-paste(wd,'/',date,'_LicenceReport.xlsx',sep="")

write.xlsx2(as.data.frame(GTC_Final),file=paste(wd,'/',date,'_LicenceReport.xlsx',sep=""),row.names = FALSE,append=TRUE,sheetName="FleetDrivers")
write.xlsx2(as.data.frame(Lite_Final),file=paste(wd,'/',date,'_LicenceReport.xlsx',sep=""),row.names = FALSE,append=TRUE,sheetName="LiteDriversVehicles")

library(RDCOMClient)

base_list<-'Daria.Alekseeva@greentomatocars.com;antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;tyrone.hunte@greentomatocars.com ;sophie.jacobsen@greentomatocars.com;Tahir.Nazir@greentomatocars.com;arti.ram@greentomatocars.com;gtcsignup@greentomatocars.com;Yinka.Ogunniyi@greentomatocars.com;Yinka.Ogunniyi@greentomatocars.com;muaaz.sarfaraz@greentomatocars.com'
#base_list<-'muaaz.sarfaraz@greentomatocars.com;'

# Send mail for 3D
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["subject"]] = 'Driver Licences and vehicles report'
outMail[["To"]] = base_list
#outMail[["To"]] = daily_list
outMail[["body"]] = "Good day. This is an automated e-mail. Drivers and vehicles available for us last week attached. Antony"
outMail[["Attachments"]]$Add(fullfile)
#outMail[["Attachments"]]$Add(paste(spreadsheets_dir,xlsx_files, sep='/'))
outMail$Send()
rm(list = c("OutApp","outMail"))
