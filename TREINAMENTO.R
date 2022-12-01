#Processamento de dados Jumppi 

#Para realizar o processamento precisamos 
#dos seguintes pacotes: expss e openxlsx
#install.packages("expss")
library(expss)
library(openxlsx)


#Por gentileze, todos usando o default  de salvamento em UTF-8

options(OutDec = ",", digits = 0)

#Iniciar do final####
#Como em nossos relatórios utilizamos valores percentuais
#com o símbolo do %, então teremos que criar um objeto 
#chamado "add_percent" que irá adicionar percentual a nossa
#planilha.
#Então olhem no sumário, cliquem e add percent, 
#selecinem até o final e rodem. Depois voltem aqui
#para retormarmos o passo a passo.

#Vamos inicialmente importar a base de dados que coletamos
#Chamando ela de base, utilizando o pacote rio e a função 
#import

base = rio::import("BANCOTREI.xlsx", which = 1)

head(base)


#Queremos analisar o que ?
#-Cidade, com base no sexo e rendimentos..

#Agora transforme todos os labels em caracteres, 
#usando essas linhas:

var.labels = as.character(names(base))

for(i in seq_along(base)){
  Hmisc::label(base[, i]) <- var.labels[i]
}

#RU####
#PROCESSAMENTO DE RESPOSTA ÚNICA

#Agora é preciso que você escolha quais variáveis 
#irão compor a sua lista de variáveis de respostas únicas, 
#anotando o número das colunas, que pode ser verificado na
#na própria base de dados
vars_ps = list(base[,c(3:4)], base[,c(5)]) #falar da 5


#Agora escolhemos o que desejamos cruzar com as variáveis 
#escolhidas anteriormente
vars_cruz = with(base, list(total(), SEXO))

#Feito isso, criamos a base, que no caso, vamos chamar de 
#RU, se referindo a resposta única
ru = base %>% 
calculate(cro_cpct(cell_vars = vars_ps, 
col_vars = vars_cruz)) %>%  
tab_sort_desc %>%  
set_caption("RU") %>% 
add_percent(digits = 0)


#criamos uma pasta de trabalho
wb = createWorkbook() 
#Adicionamos nesse documento criado a aba RU
sh = addWorksheet(wb, "RU")
#Gravamos todas tabelas em um documento
xl_write(ru, wb, sh)


#valor absoluto + percentual####

#Achado do arthur para análise dos valores absolutos também
# Absoluto = base %>% 
#   tab_cells(vars_ps) %>%
#   tab_cols(total(),vars_cruz) %>%
#   tab_stat_cases(label = "N", total_label = "") %>%
#   tab_stat_cpct(label="%", total_statistic = "w_cpct", 
#                 total_label = "") %>% 
#   tab_pivot(stat_position = "outside_rows") %>%  
#   set_caption("Absoluto")

#Argumentos que podem ser utilizados para organização 
#dos valores percentuais:
#“outside_rows”, “inside_rows”, 
#“outside_columns”, “inside_columns”


Absoluto = base %>%
  tab_cells(vars_ps) %>%
  tab_cols(total(),vars_cruz) %>%
  tab_stat_cases(label = "N", total_label = "") %>%
  tab_pivot(stat_position = "outside_rows") %>%
  set_caption("Absoluto")

sh2 = addWorksheet(wb, "Absoluto")
xl_write(Absoluto, wb, sh2)

#PESO####
base$Peso <- as.numeric(base$Peso)


ANALIPESO = base %>% 
  calculate(cro_cpct(cell_vars = vars_ps, 
                     col_vars = vars_cruz, 
                     total_statistic =  "w_cases",
                     weight = base$Peso)) %>%  
  tab_sort_desc %>%  
  set_caption("PESO")%>% 
  add_percent(digits = 0)


sh1 = addWorksheet(wb, "PESO")
xl_write(ANALIPESO, wb, sh1)

#RM####
#Agora a criação da aba com o processamento de respostas múltiplas
#Utilizamos a função mrset_p para analisar mais de uma coluna

rm = base %>% 
  calculate(cross_cpct(base, cell_vars = list(mrset_p("Q1"),
     mrset_p("Q2")),
    col_vars = vars_cruz)) %>%  
   tab_sort_desc %>%  set_caption("RM")%>% 
  add_percent(digits = 0)




sh2 = addWorksheet(wb, "RM")
xl_write(rm, wb, sh2)

#MÉDIA####

vmedias = list(base[,c(12)])


medias_cruz = with(base, list(total(), SEXO))


MEDIA = base %>% 
  calculate(cro_mean(cell_vars = vmedias, 
   col_vars = medias_cruz)) %>%  
  tab_sort_desc %>%  set_caption("MEDIA")


sh3 = addWorksheet(wb, "MEDIA")
xl_write(MEDIA, wb, sh3)


#criar arquivo xls
saveWorkbook(wb, "Output.xlsx", overwrite = TRUE)


############################################################

##### para gerar o output como porcentagem

add_percent = function(x, digits = get_expss_digits(), ...){
  UseMethod("add_percent")
}


add_percent.default = function(x, digits = get_expss_digits(), ...){
  res = formatC(x, digits = digits, format = "f")
  nas = is.na(x)
  res[nas] = ""
  res[!nas] = paste0(res[!nas], "%")
  res
}



add_percent.etable = function(x, digits = get_expss_digits(), excluded_rows = "#", ...){
  included_rows = !grepl(excluded_rows, x[[1]], perl = TRUE)
  for(i in seq_along(x)[-1]){
    if(!is.character(x[[i]])){
      x[[i]][included_rows ] = add_percent(x[[i]][included_rows])
    }
  }
  x
}

