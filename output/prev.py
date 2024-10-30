from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
import pandas as pd
from pandas import DataFrame

def previsao(df: DataFrame):
    # pegar valores das linhas anteriores (lag)
    df["Quant_lag_1"] = df["Quantidade Vendida"].shift(1)
    df["Quant_lag_2"] = df["Quantidade Vendida"].shift(2)
    df["Vendas_lag_1"] = df["Vendas Totais"].shift(1)
    df["Vendas_lag_2"] = df["Vendas Totais"].shift(2)
    df = df.dropna()
    
    X = df[["Quant_lag_1", "Quant_lag_2", "Vendas_lag_1", "Vendas_lag_2"]] # (dwvar indep
    y = df["Vendas Totais"] # var dep
    
    # treinar modelo
    modelo = LinearRegression()
    modelo.fit(X, y)
    
    # fazer previsao das vendas dos proximos meses
    df = df.assign(previsao_vendas=modelo.predict(X))
    
    return df