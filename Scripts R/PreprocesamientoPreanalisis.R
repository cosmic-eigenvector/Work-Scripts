library(dplyr)
library(haven)
library(psych)
library(DataExplorer)
library(visdat)
library(reshape2)
library(tidyr)
library(naniar)

# IO ####

# Base (MD)
df <- read_sav("./02. Raw Data/FinalGDC.sav")

# Verificar la base por olas
df %>% count(DIPWAVE_DP)

# Verificar la base por marcasy por ola 
df %>% group_by(MARCA) %>% count(DIPWAVE_DP)

#filtramos por Elektra (Marca 6)
df_marca <- df %>% filter(MARCA == 6) # %>% filter(month_der > 87) 

# Verificar la base por olas para la marca de interés
df_marca %>% count(DIPWAVE_DP)

# Tabla de validos para cada marca 
df %>% group_by(MARCA) %>% 
  summarise(M = sum(!is.na(MFI)),
            D = sum(!is.na(DFI)),
            S = sum(!is.na(SFI)),
            P = sum(!is.na(POWER)),
            AW = sum(FAMILIARITY_DP < 6))


df_marca <- df_marca %>% filter(FAMILIARITY_DP < 6)

# Validación de calidad de las variables verificamos que tengamos base correcta para todas las variables
describe_df_marca <- describe(df_marca)

# MD ####
# Preparamos el MD Dataframe, con todos los filtros a considerar (marca, periodo, aware, na's en Imagen)
MD <- df %>%
  filter(MARCA==6) %>% 
  filter(FAMILIARITY_DP<6)

attach(MD)

# IMAGE  ####
#Variables de Imagen
Image <- MD %>% select(starts_with("IMAGERY"))

describe(Image)
sapply(Image, function(x) sum(is.na(x)))
visdat::vis_dat(Image, sort_type=FALSE)
plot_missing(Image)
plot_intro(Image)
plot_bar(Image)
cor.plot(Image)
plot_correlation(na.omit(Image), type = "c")
apply(Image, 2, sum)
Image  

names(Image) <- c(
  'Get My Money In Full And In Time',
  'Quick And Easy To Collect Remittances',
  'Quick Resolution In Case Of A Problem',
  'Additional Benefits Or Loyalty Programme',
  'Trained Staff And Lean Process',
  'Extended Hours Of Operation',
  'No Limit In Amount Or Number Of Transfers',
  'Send Money Through Digital Channels',
  'Insurance Against Theft',
  'Low Requirements',
  'Safe Place',
  'Branches Near My Home Or Work'
)

describe_Imagen <- describe(Image)

# TBCA ####
#Variables de TBCA
TBCA <- MD %>% select(starts_with("TBCA")) 
plot_missing(TBCA)
describe(TBCA)

# Eliminamos los medios con <10% de respaldos, el neto TBCA y los Somewhere else y None 
TBCA <- TBCA %>% select(
  -TBCA_DER_DP,
  -TBCA_FOLLOWUP_MEDIO1,
  -TBCA_FOLLOWUP_MEDIO9,
  -TBCA_FOLLOWUP_MEDI10
)

names(TBCA) <- c(
  'Tv',
  'Radio',
  'Social Media',
  'Search Or Online Website',
  'Outdoors',
  'Friend And Family Recommendation',
  'Outside Of The Branch'
)

describe_TBCA <- describe(TBCA)

# ACT ####
# Variables de barreras y facilitadores ACT
ACT <- MD %>% select(starts_with('ACTIVATION'))
plot_missing(ACT)
describe(ACT)

names(ACT) <- c(
  'Security While Collecting Money',
  'Close To My Home',
  'Extended Hours To Collect Money',
  'Simple Documentation For Colleting Money',
  'No Queues To Collect Money',
  'Insurance Against Theft_'
)

describe_ACT <- describe(ACT)

# BRAND HEALTH KPI ####
#Variables de Marca
Filtro <- MD %>% select(AFFINITY_DP, DYNAMIC_DP, UNIQUE_DP, MEETS_NEEDS_DP, WORTH_DP, PRICE_DP, VBN_DP)
plot_missing(Filtro)
describe(Filtro)
plot_intro(Filtro)
apply(Filtro, 2, sum)

#para todas las variables de marca tomamos T2B
Filtro <- Filtro %>% 
  mutate(FairAmount = ifelse(VBN_DP < 3, 1, 0),
         Affinity = ifelse(AFFINITY_DP > 5, 1, 0),
         Unique = ifelse(UNIQUE_DP > 5, 1, 0),
         MeetNeeds = ifelse(MEETS_NEEDS_DP > 5, 1, 0),
         Dynamic = ifelse(DYNAMIC_DP > 5, 1, 0),
         PriceLow = ifelse(PRICE_DP < 4, 1, 0), # B3B para Price - aka Precio bajo 
         Worth = ifelse(WORTH_DP > 2, 1, 0)) %>% 
  select(FairAmount, Affinity, Unique, MeetNeeds, Dynamic, PriceLow, Worth)

describe_Filtro <- describe(Filtro)

# Concentracion Final ####
df.Final <- cbind(MD %>% select(.,POWER, MFI, DFI, SFI),Filtro, Image, TBCA, ACT)
library(naniar)
vis_miss(df.Final)
df.Final[is.na(df.Final)] <- 0

names(df.Final)

#Cambio de nombres por variable
colnames(df.Final)
colnames(df.Final) <- c(
  "MDF_Power","MDF_Meaningful","MDF_Different","MDF_Salient", 'MDIN_FairAmout','MDIN_Affinity', 
  'MDIN_Unique', 'MDIN_MeetNeeds', 'MDIN_Dynamic', 'MDIN_PriceLow', 'MDIN_Worth',
  gsub(" ","_",paste0("IMG_",colnames(Image))),
  gsub(" ","_",paste0("TBCA_",colnames(TBCA))),
  gsub(" ","_",paste0("ACT_",colnames(ACT)))
  )

# Añadir variable de ponderación
df.Final$Weight <- MD$weight_DP

# IO ####

colnames(df.Final)
write.csv(df.Final,"./03. Hamonized Data/MD Final.csv")
write_sav(df.Final,'./03. Hamonized Data/MD Final.sav')

# Preanalisis de Marca Foco ####
# Base Aware!!

# Correlaciones
df.Final %>% select(-Weight) %>% cor() %>% clipr::write_clip()

# Respaldos crudos
df.Final %>% 
  select(-Weight) %>% 
  summarize(across(everything(),mean)) %>% 
  clipr::write_clip()

# Respaldos ponderados
df.Final %>%
  summarize(across(everything(), ~ weighted.mean(.x, w = Weight))) %>%
  clipr::write_clip()


# Preanalisis todas las marcas ####
# Base Total Respuestas

## Preparación del dataset ####

# Preprocesamiento imagen
Image_Total <- df %>% select(starts_with("IMAGERY"))

names(Image_Total) <- c(
  'Get My Money In Full And In Time',
  'Quick And Easy To Collect Remittances',
  'Quick Resolution In Case Of A Problem',
  'Additional Benefits Or Loyalty Programme',
  'Trained Staff And Lean Process',
  'Extended Hours Of Operation',
  'No Limit In Amount Or Number Of Transfers',
  'Send Money Through Digital Channels',
  'Insurance Against Theft',
  'Low Requirements',
  'Safe Place',
  'Branches Near My Home Or Work'
)

# Preprocesamiento TBCA 
TBCA_Total <- df %>% select(starts_with("TBCA")) 

TBCA_Total <- TBCA_Total %>% select(
  -TBCA_DER_DP,
  -TBCA_FOLLOWUP_MEDIO1,
  -TBCA_FOLLOWUP_MEDIO9,
  -TBCA_FOLLOWUP_MEDI10
)

names(TBCA_Total) <- c(
  'Tv',
  'Radio',
  'Social Media',
  'Search Or Online Website',
  'Outdoors',
  'Friend And Family Recommendation',
  'Outside Of The Branch'
)

# Preprocesamiento Activadores
ACT_Total <- df %>% select(starts_with('ACTIVATION'))

names(ACT_Total) <- c(
  'Security While Collecting Money',
  'Close To My Home',
  'Extended Hours To Collect Money',
  'Simple Documentation For Colleting Money',
  'No Queues To Collect Money',
  'Insurance Against Theft_'
)

# Preprocesamiento KPIs Salud de marca 
Filtro_Total <- df %>% select(AFFINITY_DP, DYNAMIC_DP, UNIQUE_DP, MEETS_NEEDS_DP, WORTH_DP, PRICE_DP, VBN_DP)

Filtro_Total <- Filtro_Total %>% 
  mutate(FairAmount = ifelse(VBN_DP < 3, 1, 0),
         Affinity = ifelse(AFFINITY_DP > 5, 1, 0),
         Unique = ifelse(UNIQUE_DP > 5, 1, 0),
         MeetNeeds = ifelse(MEETS_NEEDS_DP > 5, 1, 0),
         Dynamic = ifelse(DYNAMIC_DP > 5, 1, 0),
         PriceLow = ifelse(PRICE_DP < 4, 1, 0), # B3B para Price - aka Precio bajo 
         Worth = ifelse(WORTH_DP > 2, 1, 0)) %>% 
  select(FairAmount, Affinity, Unique, MeetNeeds, Dynamic, PriceLow, Worth)

# Unión final del dataset 
df.Final_Total <- cbind(df %>% select(.,POWER, MFI, DFI, SFI),Filtro_Total, Image_Total, TBCA_Total, ACT_Total)
df.Final_Total[is.na(df.Final_Total)] <- 0

colnames(df.Final_Total) <- c(
  "MDF_Power","MDF_Meaningful","MDF_Different","MDF_Salient", 'MDIN_FairAmout','MDIN_Affinity', 
  'MDIN_Unique', 'MDIN_MeetNeeds', 'MDIN_Dynamic', 'MDIN_PriceLow', 'MDIN_Worth',
  gsub(" ","_",paste0("IMG_",colnames(Image_Total))),
  gsub(" ","_",paste0("TBCA_",colnames(TBCA_Total))),
  gsub(" ","_",paste0("ACT_",colnames(ACT_Total)))
)

df.Final_Total$Weight <- df$weight_DP
df.Final_Total$Brand <- df$MARCA

## Xtabs ####

# Respaldos crudos
df.Final_Total %>% 
  select(-Weight) %>% 
  group_by(Brand) %>% 
  summarize(across(everything(),mean)) %>% 
  clipr::write_clip()

# Respaldos ponderados
df.Final_Total %>%
  group_by(Brand) %>%
  summarize(across(everything(), ~ weighted.mean(.x, w = Weight))) %>%
  clipr::write_clip()
