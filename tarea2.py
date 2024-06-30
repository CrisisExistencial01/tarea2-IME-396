import numpy as np
import pandas as pd
import scipy.stats as stats
# CONDICIONES INICIALES:
#
# Nota: se necesitan las siguientes librerias:
# pandas, scipy, numpy
excel_file = 'Datos.xlsx'
xls = pd.ExcelFile(excel_file)
identificador = 18 
colors = {
        "green": "\033[92m",
        "reset": "\033[0m", 
        "red": "\033[91m",
        "blue": "\033[94m",
        "purple": "\033[95m",
        "cyan": "\033[96m",
        "orange": "\033[93m",
        }

def prettier(text, color): #color es un string
    print(f"{colors[color]}{text}{colors['reset']}")

def getSheets():
    return {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

def x_barra(muestra): # muestra = row
    muestra_valores = muestra.drop('Número')
    return np.mean(muestra_valores)

def findRow(df, identificador):
    return df[df['Número'] == identificador]

def varianza_muestral(muestra):
    muestra_cleaned = muestra.drop('Número')
    s2 = muestra_cleaned.var(ddof=1) # ddof son los Degrees Of Freedom o grados de libertad = n-1
    s = np.var(muestra_cleaned, ddof=1)
    return s2
def poisson_probability(inf, sup, valor_lambda):
    p_x = stats.poisson.cdf(sup - 1, valor_lambda) - stats.poisson.cdf(inf - 1, valor_lambda)
    return p_x 
# para mu
def intervaloDeConfianza_student(s2, mu, n, alpha):
    t = stats.t.ppf(1 - alpha/2, n-1)
    r = t*((s2)**0.5)
    l_inf = (mu - r)/((n-1)**0.5)
    l_sup = (mu + r)/(n**0.5)
    return {"inf": l_inf, "sup": l_sup}
# para el sigma cuadrado
def intervaloDeConfianza_ji_square(s2, n, alpha):
    chi2_inf = stats.chi2.ppf(alpha/2, n-1)
    chi2_sup = stats.chi2.ppf(1-alpha/2, n-1)
    l_inf = (n-1)*s2/chi2_sup
    l_sup = (n-1)*s2/chi2_inf
    return {"inf": l_inf, "sup": l_sup}

def intervaloDeConfianza_diferencia_medias(data1, data2, alpha):

    sp2 = ((data1["n"] - 1) * data1["s2"] + (data2["n"] - 1)*data2["s2"])/(data1["n"] + data2["n"] - 2)
    t = stats.t.ppf(1 - alpha/2, data1["n"] + data2["n"] - 2)
    margen_error = t * np.sqrt(sp2 * (1/data1["n"] + 1/data2["n"]))
    diff_means = data1["mu"] - data2["mu"]
    l_inf = diff_means - margen_error 
    l_sup = diff_means + margen_error
    return {"inf": l_inf, "sup": l_sup, "sp2": sp2}

# ejercicio 1
def resolv_e1(sheet):
    toSolve = findRow(sheet, identificador)
    lambda_estimado = x_barra(toSolve.iloc[0])
    prettier("Ejercicio 1", "cyan")
    prettier(f"lambda estimado: {lambda_estimado}", "purple")
    probabilidad = poisson_probability(1, 4, lambda_estimado) 
    prettier(f"probabilidad: {probabilidad}" , "blue")
    print()

def solve_general(sheet):
    n = 25
    toSolve = findRow(sheet, identificador)
    mu_estimado = x_barra(toSolve.iloc[0])
    s_square_estimado = varianza_muestral(toSolve.iloc[0])
    data = {
        "mu": mu_estimado,
        "s2": s_square_estimado,
        "t-student": intervaloDeConfianza_student(s_square_estimado, mu_estimado, n, alpha=0.01),
        "ji-square": intervaloDeConfianza_ji_square(s_square_estimado, n, alpha=0.05),
        "t-student-005": intervaloDeConfianza_student(s_square_estimado, mu_estimado, n, alpha=0.05),

        "n": n,
    }
    return data
    
# ejercicio 2
def resolv_e2(sheet): # sirve para el 2 y el 3, hacen lo mismo
    mu_estimado = solve_general(sheet)["mu"]
    s_square_estimado = solve_general(sheet)["s2"]
    prettier(f"Estimacion μ: {mu_estimado}", "purple")
    prettier(f"Estimacion σ²: {s_square_estimado}", "blue")
    intervalo = solve_general(sheet)["t-student"]

    #mu
    prettier(f"Intervalo de confianza para mu con alpha = 0.01 y 25-1 GDL: [{intervalo['inf']}, {intervalo['sup']}]", "purple")
    # sigma cuadrado
    intervalo = solve_general(sheet)["ji-square"]
    prettier(f"Intervalo de confianza para σ² y α = 0.05  con 25-1 GDL: [{intervalo['inf']}, {intervalo['sup']}]", "purple")
    print()

# ejercicio3
def resolv_e3(sheet):
    ejercicio2 = sheets['Ejercicio 2']
    ejercicio3 = sheets['Ejercicio 3']

    mu_estimado = solve_general(sheet)["mu"]
    s_square_estimado = solve_general(sheet)["s2"]
    prettier(f"Estimacion μ: {mu_estimado}", "purple")
    prettier(f"Estimacion σ²: {s_square_estimado}", "blue")

    #mu
    intervalo = solve_general(sheet)["t-student-005"]
    prettier(f"Intervalo de confianza para mu con alpha = 0.05 y 25-1 GDL: [{intervalo['inf']}, {intervalo['sup']}]", "purple")

    intervalo = solve_general(sheet)["ji-square"]
    prettier(f"Intervalo de confianza para σ² y α = 0.05  con 25-1 GDL: [{intervalo['inf']}, {intervalo['sup']}]", "purple")

    resolv_diferencia_medias(ejercicio2, ejercicio3)

def resolv_diferencia_medias(sheet1, sheet2):
    data1 = solve_general(sheet1)
    data2 = solve_general(sheet2)
    confianza_diff_medias = intervaloDeConfianza_diferencia_medias(data2, data1, alpha=0.01)
    print()
    prettier("Intervalo de Confianza para la Diferencia de Medias: ", "cyan")
    print()
    prettier(f"Diferencia de Medias estimada: {data2['mu'] - data1['mu']}", "purple")
    prettier(f"mu1: {data1['mu']}", "green")
    prettier(f"mu2: {data2['mu']}", "purple")
    prettier(f"varianza combinada: {confianza_diff_medias['sp2']}", "green")
    prettier(f"Intervalo de confianza para la diferencia de medias para alpha = 0.01: [{confianza_diff_medias['inf']}, {confianza_diff_medias['sup']}]", "blue")
    print()
print("\n")

sheets = getSheets()
ejercicio1 = sheets['Ejercicio 1']
ejercicio2 = sheets['Ejercicio 2']
ejercicio3 = sheets['Ejercicio 3']

resolv_e1(ejercicio1)
prettier("Ejercicio 2", "cyan")
resolv_e2(ejercicio2)
prettier("Ejercicio 3", "cyan")
resolv_e3(ejercicio3)

