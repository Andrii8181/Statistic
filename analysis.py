import pandas as pd
import numpy as np
from scipy.stats import shapiro, pearsonr, spearmanr
import statsmodels.api as sm
from statsmodels.formula.api import ols
import pingouin as pg

def check_normality(data):
    stat, p = shapiro(data)
    return p  # P-value

def one_way_anova(df, factor, value):
    model = ols(f"{value} ~ C({factor})", data=df).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    return anova_table

def two_way_anova(df, factor1, factor2, value):
    model = ols(f"{value} ~ C({factor1}) + C({factor2}) + C({factor1}):C({factor2})", data=df).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    return anova_table

def three_way_anova(df, f1, f2, f3, value):
    model = ols(f"{value} ~ C({f1})*C({f2})*C({f3})", data=df).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    return anova_table

def correlation(df, col1, col2, method="pearson"):
    if method == "pearson":
        return pearsonr(df[col1], df[col2])
    else:
        return spearmanr(df[col1], df[col2])

def regression(df, x, y):
    model = ols(f"{y} ~ {x}", data=df).fit()
    return model.summary()

def effect_size_eta_squared(df, factor, value):
    aov = pg.anova(data=df, dv=value, between=factor, detailed=True)
    return aov[["Source", "np2"]]
