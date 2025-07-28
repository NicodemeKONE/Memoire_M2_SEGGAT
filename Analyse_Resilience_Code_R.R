# =============================================================================
# ANALYSE DE LA RÉSILIENCE ÉCONOMIQUE DES PAYS OCDE
# Version enrichie avec documentation et améliorations
# =============================================================================

# Installation des packages nécessaires (décommenter si besoin)
# install.packages(c('zoo', 'mFilter', 'openxlsx', 'plyr', 'ggplot2', 
#                   'som', 'kohonen', 'rpart', 'rpart.plot', 'readxl',
#                   'dplyr', 'tidyr', 'corrplot', 'randomForest'), 
#                 repos="http://cran.univ-lyon1.fr")

# Chargement des libraries
library(zoo)
library(mFilter)
library(openxlsx)
library(plyr)
library(ggplot2)
library(som)
library(kohonen)
library(rpart)
library(rpart.plot)
library(readxl)
library(dplyr)
library(tidyr)
library(corrplot)
library(randomForest)

# Configuration de l'environnement
wd1 <- getwd()
setwd(wd1)
set.seed(2021)  # Pour la reproductibilité

# =============================================================================
# ÉTAPE 1 : LECTURE ET PRÉPARATION DES DONNÉES DE PRODUCTION INDUSTRIELLE
# =============================================================================

cat("=== ÉTAPE 1 : Analyse des données de production industrielle ===\n")

# Lecture des données OCDE
prodind <- read.xlsx("prodind_ocde.xlsx", sheet = "series")
cat("Données de production chargées pour", ncol(prodind)-1, "pays\n")
cat("Période:", range(prodind[,1]), "\n")

# Conversion en série temporelle
ts.prod <- ts(prodind[,-1], start = c(1990,1), frequency = 4)

# Fonction améliorée pour calculer les moyennes en gérant les NA
myColMeans <- function(x) {
  sapply(1:ncol(x), function(i) mean(na.omit(x[,i])))
}

# Fonction améliorée pour calculer les écarts-types
myColSDs <- function(x) {
  sapply(1:ncol(x), function(i) sd(na.omit(x[,i])))
}

# =============================================================================
# CALCUL DES OUTPUT GAPS AVEC FILTRE HODRICK-PRESCOTT
# =============================================================================

cat("\n=== Calcul des output gaps ===\n")

dbogap <- NULL
dbgdpr <- NULL
crisis_year <- c(2020, 1)  # Année de crise de référence

# Boucle pour chaque pays
for(i in 1:ncol(ts.prod)) {
  country_name <- colnames(ts.prod)[i]
  x <- ts.prod[,i]
  
  # Vérification de la longueur minimale des données
  if(length(na.omit(x)) > 36) {
    cat("Traitement:", country_name, "- Observations:", length(na.omit(x)), "\n")
    
    # Interpolation des valeurs manquantes
    x <- na.approx(x, na.rm = FALSE)
    
    # Application du filtre Hodrick-Prescott (fréquence 1600 pour données trimestrielles)
    hp_result <- hpfilter(na.omit(x), freq = 1600)
    pot <- hp_result$trend
    
    # Calcul de l'output gap en %
    ogap <- log(na.omit(x) / pot) * 100
    
    # Création du profil temporel (8 trimestres avant, 7 après la crise)
    profile <- window(ogap, crisis_year - c(0,8), crisis_year + c(0,7), extend = TRUE)
    
    # Stockage des résultats
    cur <- cbind(country_name, t(as.vector(profile)))
    dbogap <- rbind(dbogap, cur)
  } else {
    cat("Données insuffisantes pour:", country_name, "\n")
  }
}

# Sauvegarde des résultats
write.table(dbogap, file = "resilience.csv", sep = ";", row.names = FALSE)
cat("Output gaps sauvegardés dans resilience.csv\n")

# =============================================================================
# CALCUL DES PROFILS MOYENS ET DE DISPERSION
# =============================================================================

cat("\n=== Calcul des profils statistiques ===\n")

# Conversion en matrice numérique
dbfull <- na.omit(dbogap[,-1])
dbfull <- apply(dbfull, 2, as.numeric)
n_countries <- nrow(dbfull)
n_periods <- ncol(dbfull)

cat("Matrice finale:", n_countries, "pays x", n_periods, "périodes\n")

# Calcul du profil moyen
profil.moyen <- colMeans(dbfull, na.rm = TRUE)
write.table(profil.moyen, "resilience_moyennes_ogap.csv", sep = ";")

# Calcul de la dispersion
profil.dispersion <- apply(dbfull, 2, sd, na.rm = TRUE)
write.table(profil.dispersion, "resilience_dispersions_ogap.csv", sep = ";")

# Création des bandes de confiance
p.sup <- profil.moyen + 2 * profil.dispersion
p.inf <- profil.moyen - 2 * profil.dispersion
prof <- cbind(p.moyen = profil.moyen, p.inf = p.inf, p.sup = p.sup)

# Graphique du profil moyen avec bandes de confiance
png("profil_moyen_ogap.png", width = 800, height = 600)
ts.plot(prof, col = c("black", "red", "red"), lty = c(1, 2, 2), 
        main = "Profil moyen des output gaps avec bandes de confiance",
        xlab = "Périodes (t-8 à t+7)", ylab = "Output gap (%)")
legend("topright", c("Moyenne", "IC 95%"), col = c("black", "red"), lty = c(1, 2))
abline(v = 9, col = "blue", lty = 3)  # Ligne de crise
text(9, max(prof), "Crise", pos = 4, col = "blue")
dev.off()

# =============================================================================
# ÉTAPE 2 : CLASSIFICATION DES PROFILS DE RÉSILIENCE (SOM)
# =============================================================================

cat("\n=== ÉTAPE 2 : Classification par cartes auto-organisatrices ===\n")

# Préparation des données
dbf <- na.omit(dbogap)
dbf.pays <- dbf[,1]
cat("Pays analysés:", length(dbf.pays), "\n")

# Application de la SOM (Self-Organizing Map)
xs <- som(dbfull, xdim = 2, ydim = 3, topol = "rect")

# Visualisation de la SOM
png("som_visualization.png", width = 800, height = 600)
plot(xs, ylim = c(-22, 17), main = "Classification SOM des profils de résilience")
dev.off()

# Sauvegarde des codes SOM
write.table(cbind(dbf.pays, xs$code.sum), "resilience_SOM_codes.csv", sep = ";")

# =============================================================================
# ANALYSE ET CATÉGORISATION DES PROFILS DE RÉSILIENCE
# =============================================================================

cat("\n=== Catégorisation des profils de résilience ===\n")

# Création du mapping
mapping <- cbind(xs$visual, pays = dbf.pays)

# Classification manuelle basée sur les positions SOM
# W = Weak rebound (rebond faible)
# P = Partial rebound (rebond partiel)  
# S = Swift rebound (rebond rapide)

W.shape <- mapping[(mapping$x == 0) & (mapping$y == 0), ]  # Position (0,0)
P.shape <- mapping[((mapping$x == 0) & (mapping$y == 1)) | 
                     ((mapping$x == 0) & (mapping$y == 2)), ]  # Positions (0,1) et (0,2)
S.shape <- mapping[((mapping$x == 1) & (mapping$y == 0)) | 
                     ((mapping$x == 1) & (mapping$y == 1)) | 
                     ((mapping$x == 1) & (mapping$y == 2)), ]  # Positions (1,*)

# Sauvegarde des groupes
write.table(W.shape, "Weak_rebound.csv", sep = ";", row.names = FALSE)
write.table(P.shape, "Partial_rebound.csv", sep = ";", row.names = FALSE)
write.table(S.shape, "Swift_rebound.csv", sep = ";", row.names = FALSE)

# Affichage des statistiques
cat("Répartition des pays:\n")
cat("- Rebond faible (W):", nrow(W.shape), "pays\n")
cat("- Rebond partiel (P):", nrow(P.shape), "pays\n") 
cat("- Rebond rapide (S):", nrow(S.shape), "pays\n")

# Création du dataframe de résilience
resilience <- data.frame(pays = dbf.pays, 
                         SOMGRP = factor(paste(xs$visual$x, xs$visual$y, sep = ":")),
                         rebound = rep("Z", length(dbf.pays)))

# Attribution automatique des catégories
for(i in 1:nrow(resilience)) {
  coord <- as.character(resilience$SOMGRP[i])
  if(coord == "0:0") resilience$rebound[i] <- "W"
  else if(coord %in% c("0:1", "0:2")) resilience$rebound[i] <- "P"
  else if(coord %in% c("1:0", "1:1", "1:2")) resilience$rebound[i] <- "S"
}

# Conversion en facteur
resilience$rebound <- factor(resilience$rebound, levels = c("W", "P", "S"))

write.table(resilience, "SOM_resilience_classification.csv", sep = ";", row.names = FALSE)

# Graphique de répartition
png("repartition_resilience.png", width = 600, height = 400)
barplot(table(resilience$rebound), 
        main = "Répartition des pays par type de résilience",
        xlab = "Type de rebond", ylab = "Nombre de pays",
        col = c("red", "orange", "green"),
        names.arg = c("Faible (W)", "Partiel (P)", "Rapide (S)"))
dev.off()

# =============================================================================
# CLASSIFICATION ALTERNATIVE AVEC KOHONEN
# =============================================================================

cat("\n=== Classification alternative avec le package Kohonen ===\n")

# Création d'une grille hexagonale
iris_grid <- somgrid(xdim = 5, ydim = 5, topo = "hexagonal")

# Application de la SOM Kohonen
xs_2 <- som(X = dbfull, grid = iris_grid)

# Visualisations multiples
png("kohonen_analysis.png", width = 1200, height = 400)
par(mfrow = c(1, 3))

# Graphique 1: Nombre d'observations par neurone
plot(xs_2, type = "counts", main = "Nombre d'observations par neurone")

# Graphique 2: Heatmap d'une variable
plot(xs_2, type = "property", property = getCodes(xs_2)[, 2], 
     main = "Heatmap - Variable 2")

# Graphique 3: Diagramme en éventail des codes
plot(xs_2, type = "codes", main = "Codes des neurones")

par(mfrow = c(1, 1))
dev.off()

# =============================================================================
# ÉTAPE 3 : ANALYSE DES FACTEURS MACROÉCONOMIQUES
# =============================================================================

cat("\n=== ÉTAPE 3 : Analyse des facteurs macroéconomiques ===\n")

# Lecture des données macroéconomiques
tryCatch({
  tsA_ocde <- read.csv("dtA_ocde.csv", header = TRUE, sep = ";")
}, error = function(e) {
  cat("Erreur lecture CSV, tentative Excel...\n")
  tsA_ocde <- read_excel("dtA_ocde.xlsx", sheet = "dtA_ocde")
})

cat("Données macro chargées:", nrow(tsA_ocde), "observations\n")

# Identification des variables et pays
noms.var <- unique(tsA_ocde$weo.ind)
pays <- unique(tsA_ocde$iso.code)

cat("Variables disponibles:", length(noms.var), "\n")
cat("Pays disponibles:", length(pays), "\n")

# Extraction des valeurs
tsA_val <- tsA_ocde[, 5:ncol(tsA_ocde)]

# Préparation de la matrice macroéconomique
dt.cr <- 2020
macro <- NULL

for (v in 1:length(noms.var)) {
  nb.inf <- (v - 1) * length(pays) + 1
  nb.sup <- v * length(pays)
  
  # Vérification des bornes
  if(nb.sup <= nrow(tsA_val)) {
    tser <- ts(t(tsA_val[nb.inf:nb.sup, ]), start = 1990, frequency = 1)
    
    # Fenêtre temporelle: 5 ans avant à 6 ans après la crise
    c.inf <- (dt.cr - 1990) - 4
    c.sup <- (dt.cr - 1990) + 6
    
    if(c.inf > 0 && c.sup <= ncol(tser)) {
      tsr <- t(tser[c.inf:c.sup, ])
      colnames(tsr) <- paste(noms.var[v], 2016:2026, sep = "_")
      macro <- cbind(macro, tsr)
    }
  }
}

cat("Matrice macro créée:", nrow(macro), "x", ncol(macro), "\n")

# Alignement avec les données de résilience
macro.p <- cbind(pays = pays, data.frame(macro))
pays.ref <- sort(pays)

# Correspondance avec les pays de l'analyse de résilience
s <- sapply(resilience$pays, function(p) which(macro.p$pays == p))
s <- s[!is.na(s)]

if(length(s) > 0) {
  fact.cart <- macro[s, ]
  fact.cart <- cbind(resilience[1:nrow(fact.cart), ], data.frame(fact.cart))
  
  cat("Données combinées:", nrow(fact.cart), "pays avec", ncol(fact.cart), "variables\n")
  
  # =============================================================================
  # ANALYSE PAR ARBRE DE DÉCISION (CART)
  # =============================================================================
  
  cat("\n=== Analyse par arbres de décision ===\n")
  
  # Nettoyage des données
  fact.cart.clean <- fact.cart[complete.cases(fact.cart), ]
  
  if(nrow(fact.cart.clean) > 10) {
    # Paramètres adaptatifs pour CART
    min_split <- max(3, round(nrow(fact.cart.clean) / 12))
    min_bucket <- max(2, round(nrow(fact.cart.clean) / 8))
    
    # Modèle 1: Variables principales
    var_names <- names(fact.cart.clean)
    econ_vars <- var_names[grepl("pib|emp|chom|inf|inv|exp|gov", var_names, ignore.case = TRUE)]
    
    if(length(econ_vars) > 0) {
      formula1 <- as.formula(paste("rebound ~", paste(econ_vars[1:min(9, length(econ_vars))], collapse = " + ")))
      
      res.cart1 <- rpart(formula1, data = fact.cart.clean, 
                         minsplit = min_split, minbucket = min_bucket,
                         control = rpart.control(cp = 0.01))
      
      # Évaluation du modèle
      cart.pred1 <- predict(res.cart1, fact.cart.clean, type = "class")
      cart.tab1 <- table(Observé = fact.cart.clean$rebound, Prédit = cart.pred1)
      acc1 <- round(sum(diag(cart.tab1)) / sum(cart.tab1) * 100, 1)
      
      cat("Modèle 1 - Précision:", acc1, "%\n")
      print(cart.tab1)
      
      # Visualisation
      png(paste0("cart_model1_", acc1, "pct.png"), width = 1000, height = 700)
      rpart.plot(res.cart1, type = 1, extra = 104, varlen = 0, faclen = 0,
                 main = paste("Facteurs de résilience - Modèle 1 (Précision:", acc1, "%)"))
      dev.off()
      
      # Règles de décision
      cat("\nRègles de décision - Modèle 1:\n")
      print(rpart.rules(res.cart1))
    }
    
    # =============================================================================
    # ANALYSE COMPLÉMENTAIRE PAR RANDOM FOREST
    # =============================================================================
    
    cat("\n=== Analyse par Random Forest ===\n")
    
    if(length(econ_vars) > 2) {
      # Préparation des données pour Random Forest
      rf_data <- fact.cart.clean[, c("rebound", econ_vars)]
      rf_data <- rf_data[complete.cases(rf_data), ]
      
      if(nrow(rf_data) > 5 && length(unique(rf_data$rebound)) > 1) {
        # Random Forest
        rf_model <- randomForest(rebound ~ ., data = rf_data, 
                                 ntree = 500, importance = TRUE)
        
        cat("Erreur OOB du Random Forest:", round(rf_model$err.rate[nrow(rf_model$err.rate), 1] * 100, 1), "%\n")
        
        # Importance des variables
        importance_scores <- importance(rf_model)
        
        # Graphique d'importance
        png("variable_importance.png", width = 800, height = 600)
        varImpPlot(rf_model, main = "Importance des variables (Random Forest)")
        dev.off()
        
        cat("\nTop 5 variables les plus importantes:\n")
        imp_sorted <- sort(importance_scores[, "MeanDecreaseGini"], decreasing = TRUE)
        print(head(imp_sorted, 5))
      }
    }
    
    # =============================================================================
    # ANALYSE DE CORRÉLATION
    # =============================================================================
    
    cat("\n=== Matrice de corrélation ===\n")
    
    # Sélection des variables numériques
    numeric_vars <- sapply(fact.cart.clean, is.numeric)
    cor_data <- fact.cart.clean[, numeric_vars]
    
    if(ncol(cor_data) > 1) {
      cor_matrix <- cor(cor_data, use = "complete.obs")
      
      # Graphique de corrélation
      png("correlation_matrix.png", width = 1000, height = 1000)
      corrplot(cor_matrix, method = "color", type = "upper", 
               tl.cex = 0.8, tl.col = "black", tl.srt = 45,
               title = "Matrice de corrélation des variables macroéconomiques")
      dev.off()
    }
    
  } else {
    cat("Données insuffisantes pour l'analyse CART\n")
  }
  
  # Sauvegarde finale
  write.xlsx(fact.cart, file = "facteurs_macro_complet.xlsx", rowNames = FALSE)
  
} else {
  cat("Erreur: Impossible d'aligner les données de résilience et macroéconomiques\n")
}

# =============================================================================
# RAPPORT DE SYNTHÈSE
# =============================================================================

cat("\n" , rep("=", 70), "\n")
cat("RAPPORT DE SYNTHÈSE - ANALYSE DE RÉSILIENCE ÉCONOMIQUE\n")
cat(rep("=", 70), "\n")

cat("1. Données traitées:\n")
cat("   - Pays analysés:", nrow(resilience), "\n")
cat("   - Période d'analyse: 8 trimestres avant à 7 trimestres après", dt.cr, "\n")

cat("\n2. Classification de résilience:\n")
resilience_summary <- table(resilience$rebound)
for(i in names(resilience_summary)) {
  type_name <- switch(i, 
                      "W" = "Rebond faible", 
                      "P" = "Rebond partiel", 
                      "S" = "Rebond rapide")
  cat("   -", type_name, ":", resilience_summary[i], "pays\n")
}

cat("\n3. Fichiers générés:\n")
generated_files <- c("resilience.csv", "resilience_moyennes_ogap.csv", 
                     "resilience_dispersions_ogap.csv", "SOM_resilience_classification.csv",
                     "facteurs_macro_complet.xlsx", "profil_moyen_ogap.png",
                     "som_visualization.png", "repartition_resilience.png")

for(file in generated_files) {
  if(file.exists(file)) {
    cat("   ✓", file, "\n")
  } else {
    cat("   ✗", file, "(non généré)\n")
  }
}

cat("\n4. Recommandations pour la suite:\n")
cat("   - Examiner les profils individuels par pays\n")
cat("   - Analyser l'impact des politiques fiscales et monétaires\n")
cat("   - Étendre l'analyse à d'autres crises historiques\n")
cat("   - Valider les résultats avec des méthodes alternatives\n")

cat("\nAnalyse terminée avec succès!\n")
cat(rep("=", 70), "\n")