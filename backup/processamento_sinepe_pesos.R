setwd("C:/Users/yuri-/OneDrive - INSTITUTO OLHAR - PESQUISA E INFORMACAO ESTRATEGICA LTDA/SINEPE/6.Acompanhamento de Campo")
setwd("C:/Users/Raphaella/OneDrive - INSTITUTO OLHAR - PESQUISA E INFORMACAO ESTRATEGICA LTDA/SINEPE/6.Acompanhamento de Campo")


base = rio::import("Sinepe-MG - Acompanhamento de Campo.xlsx", which = 1)

#install.packages("expss")
library(expss)
library(openxlsx)

var.labels = as.character(names(base))

for(i in seq_along(base)){
  Hmisc::label(base[, i]) <- var.labels[i]
}

#PROCESSAMENTO RU

vars_ps = list(base[,c(15:16)],base[,c(24:26)],base[,c(31:69)],
               base[,c(78:85)],base[,c(91:102)],base[,c(109)],base[,c(127:129)])

vars_cruz = with(base, list(total(),SITUACAO_SISTEMA,Q_3,Q_4,Q_69,BH_INTERIOR))

ru = base %>% 
      calculate(cro_cpct(cell_vars = vars_ps, 
                         col_vars = vars_cruz, total_statistic =  "w_cases" ,weight = base$PESO)) %>%  tab_sort_desc %>%  set_caption("RU")

wb = createWorkbook()
sh = addWorksheet(wb, "RU")
xl_write(ru, wb, sh)


#PROCESSAMENTO RM


rm = base %>% 
        calculate(cross_cpct(base, cell_vars = list(mrset_p("Q1_"),
                                                    mrset_p("Q_5_CAT_"),
                                                    mrset_p("Q_8_CAT_"),
                                                    mrset_p("Q_43_SIM_CAT_"),
                                                    mrset_p("Q_52_CAT_"),
                                                    mrset_p("Q_65_"),
                                                    mrset_p("Q_67_"),
                                                    mrset_p("Q_68_")),
                             col_vars = vars_cruz, total_statistic =  "w_cases" ,weight = base$PESO)) %>%  tab_sort_desc %>%  set_caption("RM")




sh1 = addWorksheet(wb, "RM")
xl_write(rm, wb, sh1)


#criar arquivo xls
saveWorkbook(wb, "proc_simples_pesos.xlsx", overwrite = TRUE)

