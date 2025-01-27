# ====================================================
# Padrões ocupacionais e riscos em acidentes de trabalho com material biológico no Brasil, 2015-2019
# Autor: Júli do Brasil, Letícia Raposo
# Data: 27-01-2025
# ====================================================

# Carregando pacotes necessários
# ====================================================
required_packages <- c(
  "readxl", "dplyr", "gtsummary", "tidyverse", "hrbrthemes",
  "cluster", "NbClust", "factoextra", "FactoMineR", "officer",
  "flextable", "stringdist", "pheatmap", "RColorBrewer", "ggpubr"
)

# Instalar pacotes que não estão instalados
install_missing_packages <- function(packages) {
  for (pkg in packages) {
    if (!require(pkg, character.only = TRUE)) {
      install.packages(pkg, dependencies = TRUE)
      library(pkg, character.only = TRUE)
    }
  }
}
install_missing_packages(required_packages)

# Configuração inicial do ambiente
# ====================================================
# Configurando tema para o pacote gtsummary
library(gtsummary)
theme_gtsummary_journal("jama")

# Definindo estilos para a linguagem portuguesa
theme_gtsummary_language(
  language = "pt",
  decimal.mark = ",",
  big.mark = ".",
  iqr.sep = "-",
  ci.sep = "-",
  set_theme = TRUE
)

# Configurando função de formatação de percentuais
list(
  "tbl_summary-fn:percent_fun" = function(x)
    sprintf(x * 100, fmt = '%#.1f')
) %>%
  set_gtsummary_theme()

# Configurando o diretório de trabalho
setwd("C:\\Users\\Leticia\\Google Drive\\UNIRIO\\Projetos\\2024\\TCC - Julia do Brasil\\Dados")

# Carregando os dados
# ====================================================
dados_gerais <- read_excel("Dados_Gerais.xlsx")

# Ajustando categorias de variáveis
dados_gerais$Regiao <- factor(dados_gerais$Regiao,
                              levels = c("Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul"))

dados_gerais$Tempo_Trabalho <- factor(
  dados_gerais$Tempo_Trabalho,
  levels = c("Menos de 6 meses", "6 meses a 11 meses", "1 a 2 anos", "3 a 5 anos", "6 anos ou mais")
)

# Gerando a Tabela 1
# ====================================================
tabela1 <- dados_gerais %>%
  select(
    Ano_Notificacao, Faixa_Etaria, Sexo, Raca, Escolaridade,
    Regiao, Ocupacao, Situacao_Trab, Tempo_Trabalho,
    Exposicao_Percutanea, Exposicao_Pele_Integra,
    Exposicao_Pele_Nao_Integra, Exposicao_Mucosa, Tipo_Acidente,
    Agente_Causal, Uso_Luva, Uso_Avental, Uso_Oculos, Uso_Mascara,
    Protecao_Facial, Uso_Bota, Status_Vacinal_Hepatite_B, Fonte_Exposicao
  ) %>%
  tbl_summary()

# Análise de cluster
# ====================================================
set.seed(123)  # Para reprodutibilidade

# Função para amostrar 3000 casos por ano
amostrar_por_ano <- function(dados, n = 3000) {
  dados %>%
    group_by(Ano_Notificacao) %>%
    sample_n(n) %>%
    ungroup()
}

# Aplicando a função ao conjunto de dados
dados_amostrados <- amostrar_por_ano(dados_gerais)

# Removendo variável redundante
dados_amostrados$Ano_Notificacao <- NULL

# Análise de correspondência múltipla (MCA)
res.mca <- MCA(dados_amostrados, graph = FALSE, ncp = 20)

# Scree plot
fviz_screeplot(res.mca, addlabels = TRUE)

# Visualização de variáveis
fviz_mca_var(res.mca, choice = "mca.cor", repel = TRUE, ggtheme = theme_minimal())

# Clusterização hierárquica
res.hcpc <- HCPC(res.mca, nb.clust = -1, graph = FALSE)

# Organização de resultados e exportação de tabela
# ====================================================

# Cluster 1
cluster_1 <- as.data.frame(res.hcpc$desc.var$category$`1`)
cluster_1$Variavel <- rownames(cluster_1)
cluster_1$Cluster <- 1

# Cluster 2
cluster_2 <- as.data.frame(res.hcpc$desc.var$category$`2`)
cluster_2$Variavel <- rownames(cluster_2)
cluster_2$Cluster <- 2

# Cluster 3
cluster_3 <- as.data.frame(res.hcpc$desc.var$category$`3`)
cluster_3$Variavel <- rownames(cluster_3)
cluster_3$Cluster <- 3

# Concatenar os resultados dos clusters
df_clusters <- bind_rows(cluster_1, cluster_2, cluster_3)

# Exportando em tabela
library(officer)
library(flextable)

cluster_1$`Cla/Mod` <- round(cluster_1$`Cla/Mod`, 1)
cluster_1$`Mod/Cla` <- round(cluster_1$`Mod/Cla`, 1)
cluster_1$Global <- round(cluster_1$Global, 1)
cluster_1$p.value <- round(cluster_1$p.value, 3)
cluster_1$v.test <- round(cluster_1$v.test, 2)

cluster_2$`Cla/Mod` <- round(cluster_2$`Cla/Mod`, 1)
cluster_2$`Mod/Cla` <- round(cluster_2$`Mod/Cla`, 1)
cluster_2$Global <- round(cluster_2$Global, 1)
cluster_2$p.value <- round(cluster_2$p.value, 3)
cluster_2$v.test <- round(cluster_2$v.test, 2)

cluster_3$`Cla/Mod` <- round(cluster_3$`Cla/Mod`, 1)
cluster_3$`Mod/Cla` <- round(cluster_3$`Mod/Cla`, 1)
cluster_3$Global <- round(cluster_3$Global, 1)
cluster_3$p.value <- round(cluster_3$p.value, 3)
cluster_3$v.test <- round(cluster_3$v.test, 2)

cluster_juntos1 <- left_join(cluster_1, 
                             cluster_2, 
                             by = "Variavel")
cluster_juntos <- left_join(cluster_juntos1, 
                            cluster_3, 
                            by = "Variavel")

flextable(cluster_juntos) %>%
  save_as_docx(path = "Tabela2.docx")

# Visualização de clusters com gráficos
# ====================================================

library(stringdist)

# Calcular a matriz de distância de Levenshtein
# (Compara strings para medir a similaridade com base em inserções, deleções e substituições)
dist_matrix <- stringdistmatrix(cluster_juntos$Variavel,
                                cluster_juntos$Variavel, method = "lv")

# Usar o algoritmo de ordenação hierárquica para obter a ordem das strings
# (Agrupa strings semelhantes com base nas distâncias calculadas)
hclust_result <- hclust(as.dist(dist_matrix))
cluster_juntos <- cluster_juntos[hclust_result$order,]

# Reordenar e selecionar colunas específicas do DataFrame
cluster_juntos <- cluster_juntos[,c(6,3,1,2,5,
                                    8,9,12,
                                    14,15,18)]

# Renomear as colunas para facilitar a interpretação
colnames(cluster_juntos) <- c("Categorias",
                              "Global",
                              "Cla/Mod_Cluster1",
                              "Mod/Cla_Cluster1",
                              "v.test_Cluster1",
                              "Cla/Mod_Cluster2",
                              "Mod/Cla_Cluster2",
                              "v.test_Cluster2",
                              "Cla/Mod_Cluster3",
                              "Mod/Cla_Cluster3",
                              "v.test_Cluster3")

# Carregar pacotes para visualização
library(pheatmap)
library(RColorBrewer)

# Criar uma matriz numérica com as colunas de interesse para análise
data_matrix <-
  as.matrix(cluster_juntos[,c(5,8,11)])

# Nomear as linhas da matriz com as categorias das variáveis
rownames(data_matrix) <- c(
  "Ocupação - Limpeza/Conservação",
  "Acidente - Procedimento Médico",
  "Acidente - Manuseio de Materiais",
  "Acidente - Punção",
  "Acidente - Administração de Medicação",
  "Ocupação - Enfermagem",
  "Ocupação - Auxiliar de Laboratório",
  "Ocupação - Odontologia",
  "Exposição de Mucosa - Não",
  "Exposição de Mucosa - Sim",
  "Exposição Percutânea - Sim",
  "Exposição Percutânea - Não",
  "Hepatite B - Não Vacinado",
  "Hepatite B - Vacinado",
  "Ensino Superior Incompleto",
  "Até Ensino Médio",
  "Ensino Superior Completo",
  "Trabalho Formal",
  "Trabalho Informal",
  "Tempo de Trabalho - 1 a 2 Anos",
  "Tempo de Trabalho - 6 Anos ou Mais",
  "Tempo de Trabalho - Menos de 6 Meses",
  "Tempo de Trabalho - 6 a 11 Meses",
  "Uso de Óculos - Não",
  "Uso de Óculos - Sim",
  "Uso de Máscara - Não",
  "Uso de Máscara - Sim",
  "Uso de Avental - Não",
  "Uso de Avental - Sim",
  "Uso de Bota - Sim",
  "Uso de Bota - Não",
  "Uso de Luva - Sim",
  "Uso de Luva - Não",
  "Ocupação - Farmacêuticos",
  "Ocupação - Médicos",
  "16 a 29 Anos",
  "40 a 49 Anos",
  "30 a 39 Anos",
  "Centro-Oeste",
  "Sudeste",
  "Sul",
  "Raça Não Branca",
  "Raça Branca",
  "Sexo Masculino",
  "Sexo Feminino",
  "Fonte de Exposição - Não",
  "Fonte de Exposição - Sim",
  "Agente Causal - Agulha",
  "Agente Causal - Vidro"
)

# Nomear as colunas da matriz com os clusters
colnames(data_matrix) <- c("Cluster 1",
                           "Cluster 2",
                           "Cluster 3")

# Substituir valores infinitos por valores arbitrários para evitar problemas em cálculos
data_matrix[data_matrix == Inf] <- 100
data_matrix[data_matrix == -Inf] <- -100

# Converter a matriz para um DataFrame para facilitar a manipulação
data_matrix <- as.data.frame(data_matrix)
data_matrix$Categoria <- rownames(data_matrix)

# Criar gráficos de barras para Cluster 1
c1 <- data_matrix %>%
  filter(!is.na(`Cluster 1`)) %>% 
  filter(`Cluster 1`> 0) %>% 
  mutate(
    Categoria = reorder(Categoria, `Cluster 1`),
    Correlacao = factor(case_when(`Cluster 1` > 10 ~ "Alta",
                                  `Cluster 1` > 5 ~ "Média",
                                  TRUE ~ "Baixa"),
                        levels = c("Baixa", "Média", "Alta") 
    )) %>%
  ggplot(aes(x = Categoria, y = `Cluster 1`, fill = Correlacao)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = c(
    "Baixa" = "#00ac46",
    "Média" = "#fdc500",
    "Alta" = "#dc0000"
  )) +
  labs(title = "Grupo 1",
       x = "",
       y = "v-test",
       fill = "Representatividade") +
  theme_minimal() +
  coord_flip() +
  scale_y_continuous(
    labels = function(x)
      ifelse(x == 100, "Inf", ifelse(x == -100, "-Inf", x))
  )

# Criar gráficos de barras para Cluster 2
c2 <- data_matrix %>%
  filter(!is.na(`Cluster 2`)) %>% 
  filter(`Cluster 2`> 0) %>% 
  mutate(
    Categoria = reorder(Categoria, `Cluster 2`),
    Correlacao = factor(case_when(`Cluster 2` > 10 ~ "Alta",
                                  `Cluster 2` > 5 ~ "Média",
                                  TRUE ~ "Baixa"),
                        levels = c("Baixa", "Média", "Alta") 
    )) %>%
  ggplot(aes(x = Categoria, y = `Cluster 2`, fill = Correlacao)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = c(
    "Baixa" = "#00ac46",
    "Média" = "#fdc500",
    "Alta" = "#dc0000"
  )) +
  labs(title = "Grupo 2",
       x = "",
       y = "v-test",
       fill = "Representatividade") +
  theme_minimal() +
  coord_flip() +
  scale_y_continuous(
    labels = function(x)
      ifelse(x == 100, "Inf", ifelse(x == -100, "-Inf", x))
  )

# Criar gráficos de barras para Cluster 3
c3 <- data_matrix %>%
  filter(!is.na(`Cluster 3`)) %>% 
  filter(`Cluster 3`> 0) %>% 
  mutate(
    Categoria = reorder(Categoria, `Cluster 3`),
    Correlacao = factor(case_when(`Cluster 3` > 10 ~ "Alta",
                                  `Cluster 3` > 5 ~ "Média",
                                  TRUE ~ "Baixa"),
                        levels = c("Baixa", "Média", "Alta") 
    )) %>%
  ggplot(aes(x = Categoria, y = `Cluster 3`, fill = Correlacao)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = c(
    "Baixa" = "#00ac46",
    "Média" = "#fdc500",
    "Alta" = "#dc0000"
  )) +
  labs(title = "Grupo 3",
       x = "",
       y = "v-test",
       fill = "Representatividade") +
  theme_minimal() +
  coord_flip() +
  scale_y_continuous(
    labels = function(x)
      ifelse(x == 100, "Inf", ifelse(x == -100, "-Inf", x))
  )

# Salvar os gráficos consolidados em uma única imagem
png("Figura1.png",
    units = "cm",
    width = 15,
    height = 25,
    res = 300)
library(ggpubr)
ggarrange(
  c1,
  c2,
  c3,
  ncol = 1,
  nrow = 3,
  common.legend = TRUE,
  legend = "bottom"
)
dev.off()

