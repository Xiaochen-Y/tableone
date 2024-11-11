# check and load required packages
install_and_load <- function(packages) {
  for (pkg in packages) {
    # If the package is not installed, install it first
    if (!require(pkg, character.only = TRUE)) {
      install.packages(pkg, dependencies = TRUE)
      library(pkg, character.only = TRUE)
    } else {
      # If it is already installed, simply load it
      library(pkg, character.only = TRUE)
    }
  }
}

install_and_load(c("stats", "tableone", "nortest","dplyr","openxlsx",
                   "readxl","crayon","haven"))



Table1 <- function(myVars=myVars,catVars=catVars,contVars=contVars,data=data,
                   strata=strata,addOverall=T,pDigits=3,catDigits = 1,
                   contDigits = 2,showAllLevels=T,includeNA=T,
                   type_y="Category",simpleDes=F){
  myVars=myVars
  showAllLevels = showAllLevels
  includeNA=includeNA
  pDigits=pDigits
  catDigits=catDigits
  contDigits=contDigits
  if(simpleDes==F){
    # 1. Outcome as a categorical variable
    if(type_y=="Category"){
      normal_test <- sapply(data[contVars],function(x){
        n_size <- length(x)
        if(n_size<5000){
          p_value <- shapiro.test(x)$p.value
        }
        if(n_size>=5000){
          p_value <- lillie.test(x)$p.value
        }
        return(p_value)
      })
      nonnormal_index <- which(normal_test<0.05)
      nonnormal <- names(normal_test)[nonnormal_index]
      
      tab <- CreateTableOne(vars = myVars, strata = strata , data = data, 
                            factorVars = catVars,addOverall = addOverall,
                            includeNA = includeNA)
      tab <- print(tab, nonnormal = nonnormal, formatOptions = list(big.mark = ","),
                   printToggle=T,test = T,showAllLevels = showAllLevels,
                   noSpaces =T,catDigits = catDigits,
                   contDigits=contDigits,pDigits = pDigits)
      
      strata_n <- nrow(unique(na.omit(data[strata])))
      ##  1.1 Continuous independent variable
      ### 1.1.1 Binary outcome variable
      if(strata_n==2){
        Cont_stat <- c()
        Cont_p <- c()
        for (var in contVars) {
          #normal t.test
          if(normal_test[var]>=0.05){
            cont_test <- t.test(as.formula(paste0(var,"~",strata)),
                                data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          #nonormal wilcox.test
          if(normal_test[var]<0.05){
            cont_test <- wilcox.test(as.formula(paste0(var,"~",strata)),
                                     data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          Cont_stat <- c(Cont_stat,cont_stat)
          Cont_p <- c(Cont_p,cont_p)
        }
      }
      
      ### 1.1.2 Multiclass outcome variable
      if(strata_n>2){
        Cont_stat <- c()
        Cont_p <- c()
        for (var in contVars) {
          #normal anova
          if(normal_test[var]>=0.05){
            cont_test <- aov(as.formula(paste0(var,"~",strata)),
                             data=data)
            cont_test <- summary(cont_test)
            cont_stat <- cont_test[[1]]$`F value`[1]
            cont_p <- cont_test[[1]]$`Pr(>F)`[1]
          }
          #nonormal kruskal.test
          if(normal_test[var]<0.05){
            cont_test <-  kruskal.test(as.formula(paste0(var,"~",strata)),
                                       data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          Cont_stat <- c(Cont_stat,cont_stat)
          Cont_p <- c(Cont_p,cont_p)
        }
      }
      names(Cont_stat) <- contVars
      names(Cont_p) <- contVars
      ## 1.2 Categorical independent variable
      Cat_stat <- c()
      Cat_p <- c()
      for (var in catVars) {
        mytable <- xtabs(as.formula(paste("~",var,"+",strata,sep="")),
                         data=data)
        res <- tryCatch({
          ### 1.2.1 First, attempt Chi-square test
          chi_res <- chisq.test(mytable)
          list(cat_stat=chi_res$statistic,
               cat_p=chi_res$p.value)}, warning = function(w) {
                 ### 1.2.2 If a specific warning message appears, switch to Fisher's test
                 if (grepl("Chi-squared approximation may be incorrect", 
                           conditionMessage(w))) {
                   message(paste0("Note: Switched to Fisher's test for the variable: ",
                                  "'",var,"'."))
                   fisher_res <- fisher.test(mytable)
                   list(cat_stat = NA, cat_p = fisher_res$p.value)  
                   # Chi-square value is set to null, p-value is from Fisher's test
                 } else {
                   warning(w)
                 }
               })
        cat_stat <- res$cat_stat
        cat_p <- res$cat_p
        Cat_stat <- c(Cat_stat,cat_stat)
        Cat_p <- c(Cat_p,cat_p)
      }
      ## 1.3 Combine results of categorical and continuous variables
      names(Cat_stat) <- catVars
      names(Cat_p) <- catVars
      
      Statistic <- data.frame(stat=c(Cont_stat,Cat_stat),
                              pvalue=c(Cont_p,Cat_p))
      Statistic$varname <- rownames(Statistic)
      
      
      
      row_name <- rownames(tab)
      tab <- as.data.frame(tab)
      tab <- cbind(tab,varname='')
      tab$varname[which(tab$p!="")] <- myVars
      tab$normal <- ""
      tab$normal[which(tab$varname %in% nonnormal)] <- "Non-normal"
      p_normal <- data.frame(varname=names(normal_test),
                             p_normal=normal_test)
      tab <- left_join(tab,p_normal,by=c("varname"))
      tab <- left_join(tab,Statistic,by=c("varname"))
      tab$stat <- ifelse(!is.na(tab$stat),
                         format(round(tab$stat,3),nsmall = 3),
                         tab$stat)
      tab$pvalue <- ifelse(tab$pvalue<0.001,"<0.001",
                           format(round(tab$pvalue,pDigits),nsmall = pDigits))
      tab$p_normal <- ifelse(tab$p_normal<0.001,"<0.001",
                             format(round(tab$p_normal,pDigits),nsmall = pDigits))
      tab <- cbind(Characteristics=row_name,tab)
      delete_index <- which(tab$varname==strata)
      tab <- tab[-c(delete_index:c(delete_index+(strata_n-1))),]
      index <- which(colnames(tab)=="p")
      tab <- cbind(tab[,c(1:(index-1))],
                   tab[,c((index+5),(index+6),(index+3),(index+4))])
    }
    
    # 2. Outcome as a continuous variable
    if(type_y=="Continuous"){
      result <- c()
      for (var in catVars) {
        # Determine the number of categories in the categorical independent variable
        strata_n <- nrow(unique(na.omit(data[var])))
        y_norm <- sapply(data[strata],function(x){
          n_size <- length(x)
          if(n_size<5000){
            p_value <- shapiro.test(x)$p.value
          }
          if(n_size>=5000){
            p_value <- lillie.test(x)$p.value
          }
          return(p_value>=0.05)
        })
        ## 2.1 normal distribution
        if(y_norm==T){
          temp <- data %>%
            group_by(get(var)) %>%
            summarize(mean_var =mean(get(strata),na.rm = T),
                      sd_var=sd(get(strata),na.rm=T))
          ### 2.1.1 Binary independent variable
          if(strata_n==2){
            cont_test <- t.test(as.formula(paste0(strata,"~",var)),data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          ### 2.1.2 Multiclass independent variable
          if(strata_n>2){
            cont_test <- aov(as.formula(paste0(strata,"~",var)),
                             data=data)
            cont_test <- summary(cont_test)
            cont_stat <- cont_test[[1]]$`F value`[1]
            cont_p <- cont_test[[1]]$`Pr(>F)`[1]
          }
          temp <- cbind(Characteristics=c(var,rep(NA,nrow(temp)-1)),
                        temp,
                        Statistics=c(cont_stat,rep(NA,,nrow(temp)-1)),
                        P_value=c(cont_p,rep(NA,,nrow(temp)-1)))
          temp$res <- paste(format(round(temp$mean_var,2),nsmall = 2)," (",
                            format(round(temp$sd_var,2),nsmall = 2),")",sep = "")
        }
        
        ## 2.2 non-normal distribution
        if(y_norm==F){
          temp <- data %>%
            group_by(get(var)) %>%
            summarize(
              mean_var = quantile(get(strata), probs = 0.5, na.rm = TRUE),  #median
              sd_var= paste0("[",
                             format(round(quantile(get(strata), probs = 0.25, 
                                                   na.rm = TRUE),contDigits),nsmall=contDigits),
                             ", ",
                             format(round(quantile(get(strata), probs = 0.75, 
                                                   na.rm = TRUE),contDigits),nsmall=contDigits),
                             "]")#IQR
            )
          ### 2.1.1 Binary independent variable
          if(strata_n==2){
            cont_test <- wilcox.test(as.formula(paste0(strata,"~",var)),
                                     data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          ### 2.1.2 Multiclass independent variable
          if(strata_n>2){
            cont_test <-  kruskal.test(as.formula(paste0(strata,"~",var)),
                                       data=data)
            cont_stat <- cont_test$statistic
            cont_p <- cont_test$p.value
          }
          temp <- cbind(Characteristics=c(var,rep(NA,nrow(temp)-1)),
                        temp,
                        Statistics=c(cont_stat,rep(NA,,nrow(temp)-1)),
                        P_value=c(cont_p,rep(NA,,nrow(temp)-1))
          )
          temp$res <- paste(format(round(temp$mean_var,2),nsmall = 2)," ",
                            temp$sd_var,sep = "")
        }
        
        result <- rbind(result,temp)
      }
      
      result$P_value <- ifelse(result$P_value<0.001,"<0.001",
                               format(round(result$P_value,3),nsmall = 3))
      result$Statistics <- ifelse(!is.na(result$Characteristics),
                                  format(round(result$Statistics,3),nsmall = 3),
                                  result$Statistics)
      colnames(result)[2] <- "level"
      if(y_norm==T){colnames(result)[7] <- "Mean ± SD"}
      if(y_norm==F){colnames(result)[7] <- "Median [IQR]"}
      tab <- result[,c(1,2,7,5,6)]
    }
  }
  
  if(simpleDes==T){
    if(type_y=="Category"){testVars <- contVars}
    if(type_y=="Continuous"){testVars <- c(contVars,strata)}
    normal_test <- sapply(data[testVars],function(x){
      n_size <- length(x)
      if(n_size<5000){
        p_value <- shapiro.test(x)$p.value
      }
      if(n_size>=5000){
        p_value <- lillie.test(x)$p.value
      }
      return(p_value)
    })
    nonnormal_index <- which(normal_test<0.05)
    nonnormal <- names(normal_test)[nonnormal_index]
    
    tab <- CreateTableOne(vars = myVars, data = data, 
                          factorVars = catVars,
                          includeNA = includeNA)
    tab <- print(tab, nonnormal = nonnormal, formatOptions = list(big.mark = ","),
                 showAllLevels = showAllLevels,noSpaces =T,catDigits = catDigits,
                 contDigits = contDigits,pDigits = pDigits)
    row_name <- rownames(tab)
    tab <- data.frame(cbind(Characteristics=row_name,
                      tab=tab))
  }
  cat(red("Table 1 has been created\n"))
  cat(red("Code provided by: Xiaochen Yang\n"))
  return(tab)
}
