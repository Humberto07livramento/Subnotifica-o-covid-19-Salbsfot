# Script R para replicar a metodologia de avaliação de registros de óbitos por COVID-19
# Baseado no artigo "Proposta metodológica para avaliação de registros de óbitos por COVID-19"
# Utilizando dados fictícios para 50 BAIRROS de Salvador, Brasil.
# INCORPORA 'popdemo' para influenciar a SIMULAÇÃO do fator de sub-registro.
# VERSÃO CORRIGIDA E AMPLIADA

# Instalar e carregar pacotes necessários
# Certifique-se de que estes pacotes estão instalados. Se não estiverem, remova o '#' e execute.
# install.packages("dplyr")
# install.packages("tidyr")
# install.packages("popdemo") # Para simulação de dinâmica populacional (crescimento lambda)
# install.packages("writexl") # Pacote para escrever arquivos Excel (.xlsx)

library(dplyr)
library(tidyr)
library(popdemo) # Usado para simular lambda (taxa de crescimento)
library(writexl) # Necessário para exportar para Excel

# --- Notas Importantes ---
# 1. BGC vs. popdemo: Este script usa 'popdemo' para SIMULAR a taxa de crescimento populacional (lambda)
#    e usar isso para INFLUENCIAR o fator de sub-registro 'f'. ISSO NÃO É o método BGC de Brass.
#    BGC é uma técnica demográfica diferente. 'popdemo' usa Modelos Matriciais de População.
# 2. Simulação Conceitual: A ligação entre lambda simulado e 'f' é puramente conceitual para
#    esta simulação, assumindo que dinâmicas de crescimento podem se correlacionar (de forma
#    simplificada) com desafios no registro.
# 3. Dados Fictícios: Todas as taxas (sobrevivência, fertilidade) e dados iniciais são fictícios.

# --- 1. Criação do Banco de Dados Fictício ---

# Definir 50 bairros fictícios em Salvador (Análise ampliada 10x)
n_bairros <- 50
bairros <- paste0("Bairro ", 1:n_bairros)

# Definir grupos etários (5 grupos = matriz 5x5 para popdemo)
grupos_etarios <- c("0-19", "20-39", "40-59", "60-79", "80+")
n_stages <- length(grupos_etarios)

# Criar um data frame para a população fictícia por bairro, grupo etário e sexo
set.seed(456) # para reprodutibilidade
dados_populacao <- expand.grid(
  Bairro = bairros,
  Grupo_Etario = grupos_etarios,
  Sexo = c("Masculino", "Feminino")
) %>%
  mutate(Populacao = sample(500:25000, n(), replace = TRUE)) # Aumentar um pouco o range da população

# Criar um data frame para os óbitos observados fictícios por bairro, grupo etario, sexo e causa
# Causas: COVID-19, Mal Definidas (CMD), Códigos Garbage (CG), Outras Causas
dados_obitos_observados <- expand.grid(
  Bairro = bairros,
  Grupo_Etario = grupos_etarios,
  Sexo = c("Masculino", "Feminino"),
  Causa = c("COVID-19", "Mal Definidas", "Codigos Garbage", "Outras Causas")
) %>%
  # Ajustar a amostragem para ter mais chance de óbitos > 0
  mutate(Obitos_Observados = case_when(
    Causa == "COVID-19" ~ sample(0:50, n(), replace = TRUE, prob = c(0.4, rep(0.6/50, 50))),
    Causa == "Mal Definidas" ~ sample(0:30, n(), replace = TRUE, prob = c(0.5, rep(0.5/30, 30))),
    Causa == "Codigos Garbage" ~ sample(0:40, n(), replace = TRUE, prob = c(0.5, rep(0.5/40, 40))),
    Causa == "Outras Causas" ~ sample(0:150, n(), replace = TRUE, prob = c(0.2, rep(0.8/150, 150))),
    TRUE ~ sample(0:10, n(), replace = TRUE)
  ))

# Juntar os data frames e calcular o total de óbitos observados por bairro, sexo e causa
dados_ficticios_brutos <- dados_populacao %>%
  left_join(dados_obitos_observados, by = c("Bairro", "Grupo_Etario", "Sexo")) %>%
  # CORREÇÃO: Lidar com possíveis NAs se o join não for perfeito (não deve acontecer com expand.grid)
  filter(!is.na(Causa)) %>%
  group_by(Bairro, Sexo, Causa) %>%
  summarise(Obitos_Observados_Total = sum(Obitos_Observados, na.rm = TRUE), .groups = 'drop') %>%
  # Calcular total por bairro/sexo para cálculo de proporções depois
  group_by(Bairro, Sexo) %>%
  mutate(Total_Obitos_Bairro_Sexo = sum(Obitos_Observados_Total, na.rm = TRUE)) %>%
  ungroup()


# --- 2. Replicação Conceitual da Metodologia ---

# --- Passo 1: Estimação (Simulação) do Sub-registro de Óbitos via popdemo ---

# --- 1a. Simular Matrizes de Projeção Populacional (MPMs) e Calcular Lambda ---
cat("Iniciando simulação demográfica (popdemo)...\n")
set.seed(555)
simulacao_demografica <- expand.grid(Bairro = bairros, Sexo = c("Masculino", "Feminino")) %>%
  rowwise() %>%
  mutate(
    # Simular taxas de sobrevivência (P_i)
    survival_rates = list(runif(n_stages,
                                # Bairros ímpares terão taxas ligeiramente menores (simulação de desvantagem)
                                min = ifelse(Sexo == "Feminino", 0.88, 0.82) - ifelse(as.numeric(gsub("Bairro ", "", Bairro)) %% 2 != 0, 0.03, 0),
                                max = ifelse(Sexo == "Feminino", 0.99, 0.97) - ifelse(as.numeric(gsub("Bairro ", "", Bairro)) %% 2 != 0, 0.02, 0))),
    # Simular taxas de fertilidade (F_i)
    fertility_rates = list(c(0,
                             runif(1, 0.1, 0.6) * ifelse(Sexo == "Feminino", 1, 0), # Estágio 2 (20-39)
                             runif(1, 0.01, 0.2)* ifelse(Sexo == "Feminino", 1, 0), # Estágio 3 (40-59)
                             0, 0))
  ) %>%
  # Construir a matriz de projeção (A)
  mutate(
    matrix_A = list({
      A <- matrix(0, nrow = n_stages, ncol = n_stages)
      A[1, ] <- fertility_rates
      diag(A[2:n_stages, 1:(n_stages-1)]) <- survival_rates[1:(n_stages-1)]
      A[n_stages, n_stages] <- survival_rates[n_stages]
      A
    })
  ) %>%
  # Calcular lambda usando popdemo::eigs
  mutate(
    # Adicionado tryCatch para lidar com matrizes que podem não convergir (raro aqui, mas boa prática)
    lambda = tryCatch({
      Re(popdemo::eigs(matrix_A, what = "lambda"))
    }, error = function(e) {
      warning(paste("Erro no cálculo de lambda para:", Bairro, Sexo, "-", e$message))
      NA_real_ # Retornar NA em caso de erro
    })
  ) %>%
  ungroup() %>%
  # CORREÇÃO: Usar dplyr::select para selecionar colunas após o pipe
  dplyr::select(Bairro, Sexo, lambda) %>%
  filter(!is.na(lambda)) # Remover linhas onde lambda falhou (se houver)

cat("Simulação demográfica concluída.\n")
print("Amostra da Taxa de Crescimento Populacional (Lambda) Simulada:")
print(head(simulacao_demografica))
cat("\n")

# --- 1b. Usar Lambda Simulado para Influenciar o Fator de Sub-registro (f) ---
set.seed(789)
# CORREÇÃO: Certificar-se de que simulacao_demografica foi criada
if (!exists("simulacao_demografica") || nrow(simulacao_demografica) == 0) {
  stop("Erro: Objeto 'simulacao_demografica' não foi criado ou está vazio. Verifique a seção 1a.")
}

fatores_subregistro_ficticios <- simulacao_demografica %>%
  mutate(
    Fator_Subregistro = case_when(
      lambda > 1.015 ~ runif(n(), 1.15, 1.50),
      lambda >= 0.995 ~ runif(n(), 1.05, 1.25),
      lambda < 0.995 ~ runif(n(), 1.02, 1.20),
      TRUE ~ runif(n(), 1.05, 1.20)
    )
  )

# Juntar fatores lambda E f aos dados brutos agregados
dados_com_fatores <- dados_ficticios_brutos %>%
  left_join(fatores_subregistro_ficticios, by = c("Bairro", "Sexo"))

# Verificar se o join funcionou (se há NAs em Fator_Subregistro ou lambda)
if(any(is.na(dados_com_fatores$Fator_Subregistro)) || any(is.na(dados_com_fatores$lambda))) {
  warning("Alerta: Existem NAs nas colunas Fator_Subregistro ou lambda após o join. Verifique 'simulacao_demografica'.")
  # Opcional: filtrar NAs ou imputar valores médios/padrão
  # dados_com_fatores <- dados_com_fatores %>% filter(!is.na(Fator_Subregistro), !is.na(lambda))
}


print("Amostra dos Fatores de Sub-registro Fictícios (Influenciados por Lambda):")
print(head(fatores_subregistro_ficticios))
cat("\n")


# --- Passo 2: Redistribuição (Simulação) de Óbitos de Códigos Garbage (CG) ---
cat("Calculando redistribuição de Códigos Garbage...\n")
set.seed(101)
redistribuicao_cg_ficticia <- dados_com_fatores %>%
  filter(Causa == "Codigos Garbage") %>%
  # CORREÇÃO: Usar Fator_Subregistro já calculado (embora conceitualmente a redistribuição possa vir antes)
  # A fórmula final do artigo é O_esp^j = f * O_obs^j + CMD_j + CG_j, onde CMD_j e CG_j são as QUANTIDADES redistribuídas.
  # Então, calculamos a quantidade baseada nos óbitos OBSERVADOS.
  mutate(Proporcao_Redistribuir_CG_para_COVID = runif(n(), 0.3, 0.7)) %>%
  mutate(Redistribuido_CG_para_COVID = Obitos_Observados_Total * Proporcao_Redistribuir_CG_para_COVID) %>%
  # CORREÇÃO: Usar dplyr::select
  dplyr::select(Bairro, Sexo, Obitos_CG_Observados = Obitos_Observados_Total, Proporcao_Redistribuir_CG_para_COVID, Redistribuido_CG_para_COVID)

print("Amostra da Redistribuição de Códigos Garbage para COVID-19:")
print(head(redistribuicao_cg_ficticia))
cat("\n")

# --- Passo 3: Redistribuição (Simulação) de Óbitos de Causas Mal Definidas (CMD) ---
cat("Calculando redistribuição de Causas Mal Definidas...\n")
set.seed(112)
# Calcular Proporções
# CORREÇÃO: Usar dplyr::select corretamente
dados_proporcoes <- dados_com_fatores %>%
  filter(Causa %in% c("COVID-19", "Mal Definidas")) %>%
  # Selecionar colunas necessárias ANTES do pivot_wider para clareza
  dplyr::select(Bairro, Sexo, Causa, Obitos_Observados_Total, Total_Obitos_Bairro_Sexo) %>%
  pivot_wider(names_from = Causa, values_from = Obitos_Observados_Total, values_fill = 0) %>%
  # Verificar se as colunas existem após pivot_wider antes de renomear
  rename_with(~ "Obitos_CMD_Observados", contains("Mal Definidas")) %>%
  rename_with(~ "Obitos_COVID_Observados_Agreg", contains("COVID-19")) %>%
  mutate(
    Prop_CMD = ifelse(Total_Obitos_Bairro_Sexo == 0, 0, Obitos_CMD_Observados / Total_Obitos_Bairro_Sexo),
    Prop_COVID = ifelse(Total_Obitos_Bairro_Sexo == 0, 0, Obitos_COVID_Observados_Agreg / Total_Obitos_Bairro_Sexo)
  )

# Verificar se dados_proporcoes foi criado corretamente
if (!exists("dados_proporcoes") || nrow(dados_proporcoes) == 0) {
  stop("Erro: Objeto 'dados_proporcoes' não foi criado ou está vazio. Verifique o passo anterior.")
}
if (!"Obitos_CMD_Observados" %in% colnames(dados_proporcoes)){
  stop("Erro: Coluna 'Obitos_CMD_Observados' não encontrada em 'dados_proporcoes' após pivot/rename.")
}


# Simular Coeficiente de Redistribuição (beta do Ledermann)
set.seed(131)
coeficientes_cmd_ficticios <- dados_proporcoes %>%
  distinct(Sexo) %>%
  mutate(Coeficiente_Redistribuicao_CMD = runif(n(), 0.15, 0.45))

# Aplicar coeficiente para redistribuir
redistribuicao_cmd_ficticia <- dados_proporcoes %>%
  left_join(coeficientes_cmd_ficticios, by = "Sexo") %>%
  # CORREÇÃO: Certificar que Obitos_CMD_Observados existe
  mutate(Redistribuido_CMD_para_COVID = Obitos_CMD_Observados * Coeficiente_Redistribuicao_CMD) %>%
  # CORREÇÃO: Usar dplyr::select
  dplyr::select(Bairro, Sexo, Obitos_CMD_Observados, Coeficiente_Redistribuicao_CMD, Redistribuido_CMD_para_COVID)

print("Amostra da Redistribuição de Causas Mal Definidas para COVID-19:")
print(head(redistribuicao_cmd_ficticia))
cat("\n")

# --- Cálculo dos Óbitos Esperados por COVID-19 ---
cat("Calculando óbitos esperados por COVID-19...\n")
# Fórmula: O_esp^j = f * O_obs^j + CMD_j + CG_j

# Obter óbitos observados por COVID-19 (junto com f e lambda)
# CORREÇÃO: Usar dplyr::select corretamente
obitos_covid_observados <- dados_com_fatores %>%
  filter(Causa == "COVID-19") %>%
  # Selecionar apenas as colunas necessárias aqui
  dplyr::select(Bairro, Sexo, Obitos_COVID_Observados = Obitos_Observados_Total, Fator_Subregistro, lambda)

# Verificar se obitos_covid_observados foi criado
if (!exists("obitos_covid_observados") || nrow(obitos_covid_observados) == 0) {
  stop("Erro: Objeto 'obitos_covid_observados' não foi criado ou está vazio.")
}


# Combinar todos os componentes
# CORREÇÃO: Usar dplyr::select DENTRO do left_join para pegar colunas específicas das tabelas à direita
obitos_covid_esperados <- obitos_covid_observados %>%
  left_join(dplyr::select(redistribuicao_cmd_ficticia, Bairro, Sexo, Redistribuido_CMD_para_COVID), by = c("Bairro", "Sexo")) %>%
  left_join(dplyr::select(redistribuicao_cg_ficticia, Bairro, Sexo, Redistribuido_CG_para_COVID), by = c("Bairro", "Sexo")) %>%
  mutate(
    # Lidar com possíveis NAs dos joins (se um bairro/sexo não tinha CMD ou CG)
    Redistribuido_CMD_para_COVID = coalesce(Redistribuido_CMD_para_COVID, 0),
    Redistribuido_CG_para_COVID = coalesce(Redistribuido_CG_para_COVID, 0),
    # Aplicar fator de sub-registro AOS ÓBITOS OBSERVADOS DE COVID-19
    Obitos_COVID_Corrigidos_Subregistro = Fator_Subregistro * Obitos_COVID_Observados,
    # Calcular Óbitos Esperados
    Obitos_COVID_Esperados = Obitos_COVID_Corrigidos_Subregistro + Redistribuido_CMD_para_COVID + Redistribuido_CG_para_COVID
  ) %>%
  # Calcular diferença e incremento percentual
  mutate(
    Diferenca = Obitos_COVID_Esperados - Obitos_COVID_Observados,
    Percentual_Incremento = case_when(
      Obitos_COVID_Observados > 0 ~ (Diferenca / Obitos_COVID_Observados) * 100,
      Obitos_COVID_Observados == 0 & Obitos_COVID_Esperados > 0 ~ Inf,
      Obitos_COVID_Observados == 0 & Obitos_COVID_Esperados == 0 ~ 0,
      TRUE ~ NA_real_
    )
  )

print("Amostra do Resumo de Óbitos Observados vs. Esperados por COVID-19:")
print(head(obitos_covid_esperados %>%
             dplyr::select(Bairro, Sexo, Obitos_COVID_Observados, Obitos_COVID_Esperados, Diferenca, Percentual_Incremento)))
cat("\n")

# --- Preparar Dados para Exportação Excel ---
cat("Preparando dados para exportação Excel...\n")

# Planilha 1: Resumo Antes e Depois
resumo_excel <- obitos_covid_esperados %>%
  # CORREÇÃO: Usar dplyr::select
  dplyr::select(Bairro, Sexo, Obitos_COVID_Observados, Obitos_COVID_Esperados, Diferenca, Percentual_Incremento) %>%
  rename(
    `Bairro` = Bairro,
    `Sexo` = Sexo,
    `Óbitos COVID Observados (SIM)` = Obitos_COVID_Observados,
    `Óbitos COVID Esperados (Corrigido)` = Obitos_COVID_Esperados,
    `Diferença (Esperados - Observados)` = Diferenca,
    `Percentual de Incremento (%)` = Percentual_Incremento
  ) %>%
  mutate(across(where(is.numeric), ~round(., 3)))

# Planilha 2: Detalhes dos Cálculos e Fórmulas Simuladas
calculos_excel <- tibble(
  Passo = character(),
  Descricao = character(),
  Bairro = character(),
  Sexo = character(),
  Variavel = character(),
  Valor = double(),
  Formula_Conceitual_Simulada = character()
)

# CORREÇÃO: Adicionar lambda à tabela `fatores_subregistro_ficticios` antes do loop
fatores_subregistro_ficticios_detalhes <- fatores_subregistro_ficticios # Usar o df que já tem f e lambda

# Adicionar detalhes do Passo 1a: Simulação Lambda
# Verificar se a coluna lambda existe
if ("lambda" %in% colnames(fatores_subregistro_ficticios_detalhes)) {
  for (i in 1:nrow(fatores_subregistro_ficticios_detalhes)) {
    calculos_excel <- calculos_excel %>% add_row(
      Passo = "Passo 1a: Simulação popdemo",
      Descricao = "Taxa Crescimento Assintótico (Lambda)",
      Bairro = fatores_subregistro_ficticios_detalhes$Bairro[i],
      Sexo = fatores_subregistro_ficticios_detalhes$Sexo[i],
      Variavel = "lambda",
      Valor = fatores_subregistro_ficticios_detalhes$lambda[i],
      Formula_Conceitual_Simulada = "Calculado via popdemo::eigs(Matriz Fictícia)"
    )
  }
} else {
  cat("Aviso: Coluna 'lambda' não encontrada em 'fatores_subregistro_ficticios_detalhes'. Detalhes do Passo 1a não adicionados.\n")
}


# Adicionar detalhes do Passo 1b: Fator Sub-registro 'f'
if ("Fator_Subregistro" %in% colnames(fatores_subregistro_ficticios_detalhes)) {
  for (i in 1:nrow(fatores_subregistro_ficticios_detalhes)) {
    calculos_excel <- calculos_excel %>% add_row(
      Passo = "Passo 1b: Sub-registro",
      Descricao = "Fator de correção 'f'",
      Bairro = fatores_subregistro_ficticios_detalhes$Bairro[i],
      Sexo = fatores_subregistro_ficticios_detalhes$Sexo[i],
      Variavel = "Fator_Subregistro (f)",
      Valor = fatores_subregistro_ficticios_detalhes$Fator_Subregistro[i],
      Formula_Conceitual_Simulada = "Simulado (runif) baseado em faixas de lambda (NÃO É BGC)"
    )
  }
} else {
  cat("Aviso: Coluna 'Fator_Subregistro' não encontrada. Detalhes do Passo 1b não adicionados.\n")
}

# Adicionar componente do cálculo final relacionado ao Passo 1
if ("Obitos_COVID_Corrigidos_Subregistro" %in% colnames(obitos_covid_esperados)) {
  for (i in 1:nrow(obitos_covid_esperados)) {
    calculos_excel <- calculos_excel %>% add_row(
      Passo = "Cálculo Final (Componente)",
      Descricao = "Óbitos COVID Obs. corrigidos por 'f'",
      Bairro = obitos_covid_esperados$Bairro[i],
      Sexo = obitos_covid_esperados$Sexo[i],
      Variavel = "f * O_obs^j",
      Valor = obitos_covid_esperados$Obitos_COVID_Corrigidos_Subregistro[i],
      Formula_Conceitual_Simulada = "Fator_Subregistro * Óbitos_COVID_Observados"
    )
  }
} else {
  cat("Aviso: Coluna 'Obitos_COVID_Corrigidos_Subregistro' não encontrada.\n")
}

# Adicionar detalhes do Passo 2: Redistribuição de CG
if (exists("redistribuicao_cg_ficticia") && nrow(redistribuicao_cg_ficticia) > 0 && all(c("Redistribuido_CG_para_COVID", "Obitos_CG_Observados", "Proporcao_Redistribuir_CG_para_COVID") %in% colnames(redistribuicao_cg_ficticia))) {
  for (i in 1:nrow(redistribuicao_cg_ficticia)) {
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 2: Redist. CG", Descricao = "Óbitos CG Observados", Bairro = redistribuicao_cg_ficticia$Bairro[i], Sexo = redistribuicao_cg_ficticia$Sexo[i], Variavel = "CG_obs", Valor = redistribuicao_cg_ficticia$Obitos_CG_Observados[i], Formula_Conceitual_Simulada = "Soma óbitos Codigos Garbage")
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 2: Redist. CG", Descricao = "Proporção Redistribuição (Simulada)", Bairro = redistribuicao_cg_ficticia$Bairro[i], Sexo = redistribuicao_cg_ficticia$Sexo[i], Variavel = "Prop_CG_Redist", Valor = redistribuicao_cg_ficticia$Proporcao_Redistribuir_CG_para_COVID[i], Formula_Conceitual_Simulada = "Simulada (runif)")
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 2: Redist. CG", Descricao = "Óbitos CG Redistribuídos para COVID (CG_j)", Bairro = redistribuicao_cg_ficticia$Bairro[i], Sexo = redistribuicao_cg_ficticia$Sexo[i], Variavel = "CG_j", Valor = redistribuicao_cg_ficticia$Redistribuido_CG_para_COVID[i], Formula_Conceitual_Simulada = "CG_obs * Prop_CG_Redist")
  }
} else {
  cat("Aviso: Objeto 'redistribuicao_cg_ficticia' não encontrado, vazio ou faltando colunas.\n")
}

# Adicionar detalhes do Passo 3: Redistribuição de CMD
if (exists("redistribuicao_cmd_ficticia") && nrow(redistribuicao_cmd_ficticia) > 0 && all(c("Redistribuido_CMD_para_COVID", "Obitos_CMD_Observados", "Coeficiente_Redistribuicao_CMD") %in% colnames(redistribuicao_cmd_ficticia))) {
  for (i in 1:nrow(redistribuicao_cmd_ficticia)) {
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 3: Redist. CMD", Descricao = "Óbitos CMD Observados", Bairro = redistribuicao_cmd_ficticia$Bairro[i], Sexo = redistribuicao_cmd_ficticia$Sexo[i], Variavel = "CMD_obs", Valor = redistribuicao_cmd_ficticia$Obitos_CMD_Observados[i], Formula_Conceitual_Simulada = "Soma óbitos Causas Mal Definidas")
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 3: Redist. CMD", Descricao = "Coeficiente Redistribuição (Simulado)", Bairro = redistribuicao_cmd_ficticia$Bairro[i], Sexo = redistribuicao_cmd_ficticia$Sexo[i], Variavel = "Coef_CMD_Redist (beta)", Valor = redistribuicao_cmd_ficticia$Coeficiente_Redistribuicao_CMD[i], Formula_Conceitual_Simulada = "Simulado (runif - Método real: Ledermann)")
    calculos_excel <- calculos_excel %>% add_row(Passo = "Passo 3: Redist. CMD", Descricao = "Óbitos CMD Redistribuídos para COVID (CMD_j)", Bairro = redistribuicao_cmd_ficticia$Bairro[i], Sexo = redistribuicao_cmd_ficticia$Sexo[i], Variavel = "CMD_j", Valor = redistribuicao_cmd_ficticia$Redistribuido_CMD_para_COVID[i], Formula_Conceitual_Simulada = "CMD_obs * Coef_CMD_Redist")
  }
} else {
  cat("Aviso: Objeto 'redistribuicao_cmd_ficticia' não encontrado, vazio ou faltando colunas.\n")
}

# Adicionar detalhes do Cálculo Final: Óbitos Esperados
if (exists("obitos_covid_esperados") && nrow(obitos_covid_esperados) > 0 && "Obitos_COVID_Esperados" %in% colnames(obitos_covid_esperados)) {
  for (i in 1:nrow(obitos_covid_esperados)) {
    calculos_excel <- calculos_excel %>% add_row(
      Passo = "Cálculo Final",
      Descricao = "Óbitos Esperados por COVID-19 (O_esp^j)",
      Bairro = obitos_covid_esperados$Bairro[i],
      Sexo = obitos_covid_esperados$Sexo[i],
      Variavel = "O_esp^j",
      Valor = obitos_covid_esperados$Obitos_COVID_Esperados[i],
      Formula_Conceitual_Simulada = "f * O_obs^j + CMD_j + CG_j"
    )
  }
} else {
  cat("Aviso: Objeto 'obitos_covid_esperados' não encontrado, vazio ou faltando coluna.\n")
}

# Arredondar valores numéricos na planilha de cálculos
calculos_excel <- calculos_excel %>%
  mutate(Valor = round(Valor, 3))

# --- Exportar para Arquivo Excel na pasta Documentos ---
cat("Exportando resultados para Excel...\n")

# Obter o caminho para a pasta Documentos do usuário no Windows
pasta_documentos <- file.path(Sys.getenv("USERPROFILE"), "Documents")

# Verificar se a pasta Documentos existe, se não, usar o diretório de trabalho atual
if (!dir.exists(pasta_documentos)) {
  warning("Pasta 'Documents' padrão não encontrada. O arquivo será salvo no diretório de trabalho atual.")
  pasta_documentos <- getwd() # Usar diretório de trabalho atual como fallback
}

# Nome do arquivo Excel
nome_arquivo_excel <- "Avaliacao_Obitos_COVID_Salvador50_Ficticio_popdemo_CORRIGIDO.xlsx"

# Caminho completo para salvar o arquivo
caminho_completo_excel <- file.path(pasta_documentos, nome_arquivo_excel)

# Criar uma lista com os data frames a serem exportados para cada aba
if (exists("resumo_excel") && exists("calculos_excel")) {
  # Ordenar a planilha de detalhes para melhor visualização
  calculos_excel_ordenado <- calculos_excel %>%
    arrange(Bairro, Sexo, factor(Passo, levels = c("Passo 1a: Simulação popdemo", "Passo 1b: Sub-registro", "Passo 2: Redist. CG", "Passo 3: Redist. CMD", "Cálculo Final (Componente)", "Cálculo Final")))
  
  lista_planilhas <- list(
    `Resumo Antes e Depois (50 B)` = resumo_excel,
    `Detalhes Calculos (popdemo Sim)` = calculos_excel_ordenado
  )
  
  # Escrever o arquivo Excel
  tryCatch({
    write_xlsx(lista_planilhas, path = caminho_completo_excel)
    cat(paste0("\nArquivo Excel '", nome_arquivo_excel, "' gerado com sucesso na pasta:\n", pasta_documentos, "\n"))
  }, error = function(e) {
    cat(paste0("\nErro ao salvar o arquivo Excel em '", caminho_completo_excel, "'.\nVerifique as permissões ou se o arquivo está aberto.\nErro: ", e$message, "\n"))
  })
  
} else {
  cat("\nErro: Data frames para exportação Excel ('resumo_excel' ou 'calculos_excel') não foram criados ou estão incompletos. Arquivo Excel não gerado.\n")
}

cat("--- Fim do script ---\n")

