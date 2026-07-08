# setwd("C:/Users/Naranjoj/OneDrive - Kantar/Path Analysis/Arcos Dorados/Argentina/")

libraries <- c('data.table', 'haven', 'magrittr', 'stringr', 'psych', 'stringi','janitor')
invisible(lapply(libraries, library, character.only = TRUE))



# Funcs -------------------------------------------------------------------

table_lab <- function(x){
  labels_ <- data.table(x = as.character(attr(x, "labels")), 
                        label = names(attr(x, "labels")))
  table_ <- as.data.table(table(x))
  merge(labels_, table_, by = "x")
  
}


to_one_column <- function(dat_, columnGroups){
  
  dat <- copy(dat_)
  newDat <- data.table(row = 1:nrow(dat))
  
  for(i in columnGroups){
    currentGroup <- str_subset(names(dat), i)
    values <- str_extract(currentGroup, "\\d+") %>% 
      as.numeric()
    for(j in values){
      
      iter <- which(j == values)
      dat[[currentGroup[iter]]] <- ifelse(dat[[currentGroup[iter]]] == 1, j, 0)
    }
    
    newDat[[i]] <- rowSums(dat[, ..currentGroup])
  }
  newDat[, row := NULL]
  return(newDat)
}

get_names <- function(data, string_start, string_delete, string_replace){
  paste0(string_start, str_replace_all(str_replace_all(mapply(attr, data, 'label'), string_delete, ""),
                                       string_replace, "_"))
}



factor_analysis <- function(dat, nFact, rotate = "varimax", cor = 'poly', fm = "minres", 
                            removeFromLabel = "", qImpor = 0, negLoadings = T){
  
  fa_ <- psych::fa(as.data.frame(dat), nfactors = nFact, rotate = rotate,
                   cor = cor, fm = fm)
  
  psych::fa.diagram(fa_) 
  
  loadings <- if(negLoadings){
    apply(as.data.frame(matrix(fa_$loadings, ncol = nFact, nrow = ncol(dat))), 1, 
          function(x) ifelse(abs(x) >= quantile(abs(x), qImpor), x, NA)) %>%  t
  }else{
    apply(as.data.frame(matrix(fa_$loadings, ncol = nFact, nrow = ncol(dat))), 1, 
          function(x) ifelse(x >= quantile(x, qImpor), x, NA)) %>%  t
  }
  
  rownames(loadings) <- rownames(fa_$loadings)
  
  faLoadings <- data.table("varName" = rownames(loadings), loadings)
  
  groupOrder <- apply(loadings, 1, function(x) which.max(abs(x))) 
  valueOrder <- apply(loadings, 1, function(x) x[which.max(abs(x))])
  
  faLoadings <- data.table("Group" = as.numeric(groupOrder), faLoadings)
  
  faLoadings[, valueOrder := as.numeric(valueOrder)]
  
  faLoadings <- faLoadings[order(Group, -valueOrder)] %>% 
    .[, -"valueOrder"]
  
  print(paste0("Variance Explained: ", 
               round(max(fa_$Vaccounted["Cumulative Var", ]), 3)))
  
  return(list("loadings" = faLoadings, "scores" = fa_$scores, "model" = fa_))
}


label_as_colNames <- function(dat, removeFromLabel = "@"){
  newNames <- mapply(attr, dat, "label")
  
  dictNames <- data.table("Old" = names(newNames), 
                          "New" = str_to_title(janitor::make_clean_names(str_remove(newNames, removeFromLabel))))
  
  setnames(dat, 
           old = dictNames$Old, 
           new = dictNames$New)
  return(dat)
}


replace_values <- function(x, old, new){
  if(length(old) != length(new)) stop("New and Old values have different lenght")
  if(!all(unique(x) %in% old)) stop("Old values are incomplete")
  for (i in 1:length(x)) x[i] <- new[old %in% x[i]]
  return(x)
}


clean_characterVector <- function(x, prefix = "", sufix = "", 
                                  removeFromLabel = "@"){
  x <- str_remove(x, removeFromLabel) %>% 
    janitor::make_clean_names() %>% 
    str_to_title()
  
  paste0(prefix, x, sufix)
  
}

# Data --------------------------------------------------------------------

data <- read_sav('./Data/MD Final.sav') %>% 
  as.data.table() #%>% .[indice1 == 1, ] #revisar que se filtre por mcdonalds


vtable::vtable(data)


# Factor: Imagery ---------------------------------------------------------

imageryCols <- str_subset(names(data), "ACT_")

# Renaming vars 
attr(data$IMAGERY_1, "name")

imageryData <- copy(data[, ..imageryCols]) %>% 
  label_as_colNames(removeFromLabel = "ACT_")

# NA Filtering in case of needing it 
# imageryData <- imageryData %>% filter(!is.na(COLUMN))

# Frequency
imageryData %>% colMeans() %>% sort() %>% names %>% clipr::write_clip()
all(rowSums(imageryData) > 0) 

# Should we use FA?
KMO(imageryData)

# How many factors? (For reference)
scree(imageryData)
fa.parallel(imageryData, fa="fa")

#FA
FA.imagery <- factor_analysis(imageryData, nFact = 2, rotate = "quartimax", 
                              cor = "tetrachoric", fm = "ml")

# Proportion of variation in that variable explained by the n-factors
FA.imagery$model$communalities

# Loadings
FA.imagery$loadings %>% clipr::write_clip()

setnames(imageryData,
         old = names(imageryData),
         new = paste0("IMQ16_",  names(imageryData)))