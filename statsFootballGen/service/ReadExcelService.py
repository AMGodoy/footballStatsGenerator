import pandas as pd
import os


# --- URLs de ligas ---
BASE_PATH = r"C:\Users\Alonso MG\Desktop\FOLDERS\ESTADÍSTICAS POR LIGA"

ArgentinaPrimeraDiv = os.path.join(BASE_PATH, "Argentina Primera Div..xlsx")
ArgentinaPrimeraNac = os.path.join(BASE_PATH, "Argentina Primera Nac..xlsx")
BrasilSerieA = os.path.join(BASE_PATH, "Brasil Serie A.xlsx")
BrasilSerieB = os.path.join(BASE_PATH, "Brasil Serie B.xlsx")
ColombiaPrimeraA = os.path.join(BASE_PATH, "Colombia Primera A.xlsx")
EcuadorLigaPro = os.path.join(BASE_PATH, "Ecuador Liga Pro.xlsx")
FinlandVeikkausliiga = os.path.join(BASE_PATH, "Finland Veikkausliiga.xlsx")
FinlandYkkosliiga = os.path.join(BASE_PATH, "Finland Ykkosliiga.xlsx")
Iceland1Deild = os.path.join(BASE_PATH, "Iceland 1 Deild.xlsx")
IcelandUrvalsdeild = os.path.join(BASE_PATH, "Iceland Urvalsdeild.xlsx")
JaponJ1League = os.path.join(BASE_PATH, "Japon J1League.xlsx")
LigaAustralia24 = os.path.join(BASE_PATH, "Liga Australia 24.xlsx")
LigaBolivia = os.path.join(BASE_PATH, "Liga Bolivia.xlsx")
LigaChilena = os.path.join(BASE_PATH, "Liga Chilena.xlsx")
LigaEstonia = os.path.join(BASE_PATH, "Liga Estonia.xlsx")
LigaMX = os.path.join(BASE_PATH, "Liga MX.xlsx")
LigaPeruana = os.path.join(BASE_PATH, "Liga Peruana.xlsx")
LigaRumania = os.path.join(BASE_PATH, "Liga Rumania.xlsx")
LigaUruguay = os.path.join(BASE_PATH, "Liga Uruguay.xlsx")
LigaVenezuela = os.path.join(BASE_PATH, "Liga Venezuela.xlsx")
MLS = os.path.join(BASE_PATH, "MLS.xlsx")
SpainLigaHiper = os.path.join(BASE_PATH, "Spain Liga Hiper.xlsx")
SueciaAllsvenskan = os.path.join(BASE_PATH, "Suecia Allsvenskan.xlsx")
SuperLigaChina = os.path.join(BASE_PATH, "Super Liga China.xlsx")
AlemaniaBundelsiga = os.path.join(BASE_PATH, "Alemania Bundesliga.xlsx")
BelgicaLiga = os.path.join(BASE_PATH, "Belgica Liga.xlsx")
PremierLeague = os.path.join(BASE_PATH, "Premier League.xlsx")

# --- Diccionario de mapeo dinámico ---
COLUMN_MAP = {
    "Resultados": ["Equipo", "PJ", "Win%", "Draw%", "Lose%"],
    "Goles anotados": ["Equipo", "PJ", "Goles anotados", "TxG", "xG"],
    "Goles concedidos": ["Equipo", "PJ", "Goles concedidos", "TxGA", "xGA"],
    "BTTS": ["Equipo", "PJ", "BTTS", "BTTS%"],
    "Anota Primero": ["Equipo", "PJ", "Veces", "Porcentaje"],
    "Concede Primero": ["Equipo", "PJ", "Veces", "Porcentaje"],
    "Anota 2+": ["Equipo", "PJ", "Veces", "Porcentaje"],
    "Concede 2+": ["Equipo", "PJ", "Veces", "Porcentaje"],
    "Corners For": ["Equipo", "PJ", "Corners", "Promedio"],
    "Corners Against": ["Equipo", "PJ", "Corners", "Promedio"],
    "Faltas cometidas": ["Equipo", "PJ", "Faltas cometidas", "Promedio"],
    "Faltas ganadas": ["Equipo", "PJ", "Faltas ganadas", "Promedio"],
    "Tarjetas generadas": ["Equipo", "PJ", "Tarjetas generadas", "Promedio"],
    "Tarjetas en contra": ["Equipo", "PJ", "Tarjetas en contra", "Promedio"],
    "0.5 1T": ["Equipo", "PJ", "Over 0.5 1H", "Over %", "Under 0.5 1H", "Under %"],
    "0.5 2T": ["Equipo", "PJ", "Over 0.5 2H", "Over %", "Under 0.5 2H", "Under %"],
    "1.5 GOLES": ["Equipo", "PJ", "Over 1.5", "Over %", "Under 1.5", "Under %"],
    "2.5 GOLES": ["Equipo", "PJ", "Over 2.5", "Over %", "Under 2.5", "Under %"],
    "Remates a favor": ["Equipo", "PJ", "Total de tiros", "Promedio"],
    "Remates en contra": ["Equipo", "PJ", "Total de tiros concedidos", "Promedio"],
    "Remates al arco a favor": ["Equipo", "PJ", "Total de SOT", "Promedio"],
    "Remates al arco en contra": ["Equipo", "PJ", "Total de SOT concedidos", "Promedio"],
}




def leer_bloques(excel_path, hoja):
    """
    Lee una hoja con 3 tablas (Home, Away, Overall) separadas por columnas en blanco.
    """
    df = pd.read_excel(excel_path, sheet_name=hoja, header=None)
    df = df.iloc[:, 2:]  # ignoramos columnas A y B

    if hoja not in COLUMN_MAP:
        return {}

    cols = COLUMN_MAP[hoja]
    n = len(cols)

    # Home: primeras n columnas
    df_home = df.iloc[:, :n].dropna(how="all").copy()
    df_home.columns = cols

    # Away: después de la columna en blanco
    df_away = df.iloc[:, n+1:n+1+n].dropna(how="all").copy()
    df_away.columns = cols

    # Overall (no lo usamos, pero lo cargamos por si acaso)
    df_overall = df.iloc[:, 2*n+2:2*n+2+n].dropna(how="all").copy()
    df_overall.columns = cols

    return {
        "home": df_home,
        "away": df_away,
        "overall": df_overall
    }


def formatear_fila(row, rol):
    pj = row["PJ"]
    partes = [f"{row['Equipo']} ({rol}): jugó {pj} partidos"]

    for col in row.index:
        if col in ["Equipo", "PJ"]:
            continue
        valor = row[col]
        if col.lower() == "corners":
            partes.append(f"generó {valor}")
        elif col.lower() == "promedio":
            partes.append(f"con promedio {valor}")
        else:
            partes.append(f"{col}: {valor}")
    return ", ".join(partes)


def estadisticas_equipos(file_path, equipo_local, equipo_visita):
    xls = pd.ExcelFile(file_path)

    for hoja in xls.sheet_names:
        if hoja.lower() == "referee":
            continue

        bloques = leer_bloques(file_path, hoja)
        if not bloques:
            continue

        print(f"\n--- {hoja.upper()} ---")

        # Buscar Local en Home
        df_home = bloques["home"]
        row_local = df_home[df_home["Equipo"].str.contains(equipo_local, case=False, na=False)]
        if not row_local.empty:
            print(formatear_fila(row_local.iloc[0], "Local"))
        else:
            print(f"{equipo_local} (Local): No se encontró información en {hoja}")

        # Buscar Visita en Away
        df_away = bloques["away"]
        row_visita = df_away[df_away["Equipo"].str.contains(equipo_visita, case=False, na=False)]
        if not row_visita.empty:
            print(formatear_fila(row_visita.iloc[0], "Visita"))
        else:
            print(f"{equipo_visita} (Visita): No se encontró información en {hoja}")

def estadisticas_equipos_inter(file_path_local, equipo_local, file_path_visita, equipo_visita):
    """
    Para competiciones internacionales: imprime las estadísticas de Local y Visita unificadas por bloque.
    """
    xls_local = pd.ExcelFile(file_path_local)
    xls_visita = pd.ExcelFile(file_path_visita)

    # Recorremos todas las hojas que existan en ambos excels
    hojas = set(xls_local.sheet_names) & set(xls_visita.sheet_names)

    for hoja in hojas:
        if hoja.lower() == "referee":
            continue

        bloques_local = leer_bloques(file_path_local, hoja)
        bloques_visita = leer_bloques(file_path_visita, hoja)

        if not bloques_local or not bloques_visita:
            continue

        df_home = bloques_local["home"]
        df_away = bloques_visita["away"]

        row_local = df_home[df_home["Equipo"].str.contains(equipo_local, case=False, na=False)]
        row_visita = df_away[df_away["Equipo"].str.contains(equipo_visita, case=False, na=False)]

        if not row_local.empty or not row_visita.empty:
            print(f"\n--- {hoja.upper()} ---")

        if not row_local.empty:
            print(formatear_fila(row_local.iloc[0], "Local"))
        if not row_visita.empty:
            print(formatear_fila(row_visita.iloc[0], "Visita"))

def limpiar_dataframe(df):
    """
    Elimina filas vacías o que son encabezados tipo 'Home', 'Away', 'Unnamed', etc.
    """
    if df.empty:
        return df

    # Eliminar filas con NaN en la primera columna
    df = df.dropna(subset=[df.columns[0]])

    # Quitar filas basura
    df = df[~df[df.columns[0]].astype(str).str.contains("Home|Away|Unnamed|nan", case=False, na=False)]

    return df


def top_4_locales_visitas(file_path):
    """
    Lee los primeros 4 equipos locales y visitantes en cada hoja y los imprime en consola.
    """
    xls = pd.ExcelFile(file_path)

    for hoja in xls.sheet_names:
        if hoja.lower() == "referee":
            continue

        bloques = leer_bloques(file_path, hoja)
        if not bloques:
            continue

        # 🔥 limpiar antes de tomar los primeros 4
        df_home = limpiar_dataframe(bloques["home"]).head(4)
        df_away = limpiar_dataframe(bloques["away"]).head(4)

        print(f"\n--- PRIMEROS 4 ({hoja.upper()}) ---")

        if not df_home.empty:
            print("\n🏠 Locales:")
            for _, row in df_home.iterrows():
                print(formatear_fila(row, "Local"))

        if not df_away.empty:
            print("\n🚗 Visitantes:")
            for _, row in df_away.iterrows():
                print(formatear_fila(row, "Visita"))



# 🚀 Ejemplo de uso
if __name__ == "__main__":

    #Ejemplo: mostrar primeros 4 por hoja
    #top_4_locales_visitas(LigaMX)

    estadisticas_equipos(SuperLigaChina,"Changchun Yatai", "Qingdao Hainiu")
    print()
    """print("#### OTRO PARTIDO ######")
    estadisticas_equipos(SuperLigaChina,"Beijing","Dalian")
    print()
    print("#### OTRO PARTIDO ######")
    estadisticas_equipos(SuperLigaChina,"Shanghai Port","Wuhan")
    print()
    print("#### OTRO PARTIDO ######")
    estadisticas_equipos(SuperLigaChina,"Zhejiang","Chengdu")




    # Diferentes ligas
    #estadisticas_equipos_inter(excel_path, "A. Lima", excel_path2, "U. de Chile")"""



