import pandas as pd
import matplotlib.pyplot as plt
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
import os

load_dotenv()


def clean_column_names(df):
    """Nettoie les noms des colonnes du DataFrame
    pour les rendre compatibles avec PostgreSQL"""

    df.columns = [col.strip().replace(' ', '_').lower() for col in df.columns]
    return df


def load_data_from_excel(file_path):
    """Charge les feuilles Excel nécessaires en DataFrame"""

    sheets = {
        "Vente_France": pd.read_excel(file_path, sheet_name="Vente_France"),
        "Vente_Allemagne": pd.read_excel(file_path, sheet_name="Vente_Allemagne"),
        "Vente_Pologne": pd.read_excel(file_path, sheet_name="Vente_Pologne"),
        "Cout": pd.read_excel(file_path, sheet_name="Cout"),
        "Referentiel_Produit": pd.read_excel(file_path, sheet_name="Referentiel_Produit"),
    }
    return sheets


def connect_to_postgresql():
    """Connexion à la base de données PostgreSQL"""

    db_host = os.getenv("POSTGRES_HOST")
    db_database = os.getenv("POSTGRES_DB")
    db_user = os.getenv("POSTGRES_USER")
    db_password = os.getenv("POSTGRES_PASSWORD")
    db_port = os.getenv("POSTGRES_PORT")

    print(f"Connexion à PostgreSQL: host={db_host}, db={db_database},user={db_user}, port={db_port}")

    try:
        connection = f"postgresql://{db_user}:{db_password}@{db_host}:{db_port}/{db_database}"
        engine = create_engine(connection)
        conn = engine.connect()
        print("Connexion réussie.")
        return conn

    except Exception as e:
        print(f"Impossible de se connecter à la base de données : {e}")
        return None


def load_data_to_postgresql(sheets):
    """Charge le dictionnaire de DataFrames
    dans une base de données PostgreSQL"""

    # Connexion à la base de données
    conn = connect_to_postgresql()
    if conn:
        # Parcourir les feuilles et insérer les données dans la base de données
        for sheet_name, df in sheets.items():

            # Nettoyer les noms des colonnes
            df = clean_column_names(df)

            # Insérer les données dans la base de données
            try:
                df.to_sql(sheet_name, conn, if_exists='replace', index=False)
                print(f"Les données de la feuille {sheet_name} ont été insérées avec succès.")
            except Exception as e:
                print(f"Une erreur s'est produite lors de l'insertion des données : {e}")        
        return conn
    else:
        return None


def execute_sql_query(conn, query):
    """Exécute une requête SQL et retourne le résultat sous forme de DataFrame"""

    try:
        result = conn.execute(text(query))
        df = pd.DataFrame(result.fetchall(), columns=result.keys())
        print("Colonne du DataFrame : ", df.columns)
        return df

    except Exception as e:
        print(f"Une erreur s'est produit lors de l'exécution de la requête : {e}")
        return None


def generate_graph_bar_from_dataframe(ax, df, title, x_label, y_label):
    """Génère un graphique à barres à partir d'un DataFrame"""

    ax.bar(df[x_label], df[y_label], color=["skyblue", "green", "red"])
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)


def generate_grouped_bar_chart(ax, df, title, x_label, y_label):
    """Génère un graphique à barres groupées à partir d'un DataFrame"""

    bar_width = 0.35
    index = range(len(df))
    
    # Barres pour 2019 et 2020
    ax.bar(index, df['CA_2019'], width=bar_width, label='2019', color='skyblue')
    ax.bar([i + bar_width for i in index], df['CA_2020'], width=bar_width, label='2020', color='orange')
    
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.set_xticks([i + bar_width / 2 for i in index])
    ax.set_xticklabels(df['pays'])
    ax.legend()
    
    # Calculer l'augmentation
    df['Increase'] = df['CA_2020'] - df['CA_2019']
    max_increase_country = df.loc[df['Increase'].idxmax()]
    
    # Calculer le pourcentage d'augmentation
    increase_value = max_increase_country['Increase']
    increase_percentage = (increase_value / max_increase_country['CA_2019']) * 100

    # Ajouter une annotation pour le pays avec la plus forte augmentation
    ax.text(0.5, -0.3, f"Le pays avec la plus forte augmentation de CA est : {max_increase_country['pays']} - {increase_value} ({increase_percentage:.2f}%)",
            ha='center', va='center', transform=ax.transAxes, fontsize=11)


def section_3(axs, conn):
    """Génère les graphiques pour la section 3"""

    # Graphique 1 : Chiffre d'affaires par pays
    query_ca_by_country = """
        SELECT Pays, SUM(CAST("ca" AS NUMERIC)) AS "Total_CA"
        FROM (
            SELECT 'France' AS "pays", "ca" FROM "Vente_France"z
            UNION ALL
            SELECT 'Allemagne' AS "pays", "ca" FROM "Vente_Allemagne"
            UNION ALL
            SELECT 'Pologne' AS "pays", "ca" FROM "Vente_Pologne"
        )
        GROUP BY "pays"
        ORDER BY "Total_CA" DESC
        """
    df_ca = execute_sql_query(conn, query_ca_by_country)
    generate_graph_bar_from_dataframe(axs[0, 0], df_ca, "Chiffre d'affaires par Pays", "pays", "Total_CA")

    # Afficher les colonnes du DataFrame
    print(df_ca.columns)

    generate_graph_bar_from_dataframe(axs[0, 0], df_ca, "Chiffre d'affaires par Pays", "pays", "Total_CA")

    # Ajouter une annotation pour le pays avec le CA le plus élevé
    max_ca_country = df_ca.iloc[0]
    max_ca_value = max_ca_country["Total_CA"]
    total_ca_sum = df_ca["Total_CA"].sum()
    ca_percentage = (max_ca_value / total_ca_sum) * 100
    axs[0, 0].text(0.5, -0.3, f"Le pays avec le CA le plus élevé est : {max_ca_country['pays']} - {max_ca_value} ({ca_percentage:.2f}%)",
                   ha='center', va='center', transform=axs[0, 0].transAxes, fontsize=11)

    # Graphique 2 : Évolution du CA
    query_ca_evolution = """
    SELECT Pays,
              SUM(CASE WHEN Annee = 2020 THEN CA ELSE 0 END) AS "CA_2020",
              SUM(CASE WHEN Annee = 2019 THEN CA ELSE 0 END) AS "CA_2019"
    FROM (
        SELECT 'France' AS pays, Annee, Ca FROM "Vente_France"
        UNION ALL
        SELECT 'Allemagne' AS pays, Annee, Ca FROM "Vente_Allemagne"
        UNION ALL
        SELECT 'Pologne' AS pays, Annee, Ca FROM "Vente_Pologne"
    )
    GROUP BY pays
    ORDER BY pays
    """
    df_evolution = execute_sql_query(conn, query_ca_evolution)
    generate_grouped_bar_chart(axs[0, 1], df_evolution, "Évolution du Chiffre d'Affaires", "pays", "CA")


def generate_margin_bar_chart(ax, df_marge):
    """Génère un graphique à barres de la marge par produit."""

    ax.bar(df_marge['lib_produit'], df_marge['marge_totale'], color='red')
    ax.set_title("Marge par produit (Tous les pays)")
    ax.set_xlabel("Produit")
    ax.set_ylabel("Marge")
    ax.tick_params(axis='x', rotation=45)

    # Trouver le produit ayant généré le plus de marge
    max_margin_product = df_marge.iloc[0]
    max_margin_value = max_margin_product['marge_totale']
    total_margin = df_marge['marge_totale'].sum()
    margin_percentage = (max_margin_value / total_margin) * 100

    # Ajouter une annotation sous le graphique avec un décalage
    ax.text(0.5, -0.3, 
            f"Le produit avec la marge la plus élevée est : {max_margin_product['lib_produit']} - {max_margin_value} ({margin_percentage:.2f}%)",
            ha='center', va='center', transform=ax.transAxes, fontsize=10)


def generate_margin_distribution_pie_chart(ax, total_marges, product_name, labels):
    """Génère un graphique en camembert de la répartition de la marge par pays."""
    
    # Tracer le graphique en camembert
    ax.pie(total_marges, labels=labels, autopct='%1.1f%%', startangle=90)
    ax.set_title(f"Répartition de la marge pour : {product_name}")

    # Trouver le pays avec la plus grande part de marge
    max_idx = total_marges.argmax()  # Trouver l'indice de la marge maximale
    max_margin_country = labels[max_idx]   # Accéder au label correspondant
    max_country_margin = total_marges[max_idx]
    total_margin = total_marges.sum()
    margin_percentage = (max_country_margin / total_margin) * 100

    # Ajouter une annotation en dessous du graphique
    ax.text(0, -1.2, 
            f"Le pays avec la plus grande part de marge : {max_margin_country} - {max_country_margin} ({margin_percentage:.2f}%)",
            ha='center', va='center', fontsize=10)


def section_4(axs, conn):
    """Génère les graphiques pour la section 4."""
    
    # Requête SQL pour obtenir la marge par produit pour tous les pays
    query_cost_per_product = """
    SELECT rp.Produit,
           rp.Lib_Produit,
           SUM(vf.CA - c.Cout_France) AS Marge_France,
           SUM(va.CA - c.Cout_Allemagne) AS Marge_Allemagne,
           SUM(vp.CA - c.Cout_Pologne) AS Marge_Pologne,
           (SUM(vf.CA - c.Cout_France) + SUM(va.CA - c.Cout_Allemagne) + SUM(vp.CA - c.Cout_Pologne)) AS Marge_Totale
    FROM "Referentiel_Produit" rp
    LEFT JOIN (SELECT Produit, CA FROM "Vente_France") vf ON rp.Produit = vf.Produit
    LEFT JOIN (SELECT Produit, CA FROM "Vente_Allemagne") va ON rp.Produit = va.Produit
    LEFT JOIN (SELECT Produit, CA FROM "Vente_Pologne") vp ON rp.Produit = vp.Produit
    JOIN "Cout" c ON rp.Produit = c.Produit
    GROUP BY rp.Produit, rp.Lib_Produit
    ORDER BY Marge_Totale DESC
    """

    df_marge = execute_sql_query(conn, query_cost_per_product)

    # Graphique 3 de la marge par produit
    generate_margin_bar_chart(axs[1, 0], df_marge)

    # Trouver le produit ayant généré le plus de marge
    max_margin_product = df_marge.iloc[0]

    # Section 4 : Répartition de la marge par pays
    query_margin_distribution = f"""
    SELECT 
        SUM(vf.CA - c.Cout_France) AS Marge_France,
        SUM(va.CA - c.Cout_Allemagne) AS Marge_Allemagne,
        SUM(vp.CA - c.Cout_Pologne) AS Marge_Pologne
    FROM "Cout" c
    LEFT JOIN "Vente_France" vf ON vf.Produit = c.Produit
    LEFT JOIN "Vente_Allemagne" va ON va.Produit = c.Produit
    LEFT JOIN "Vente_Pologne" vp ON vp.Produit = c.Produit
    WHERE vf.Produit = '{max_margin_product['produit']}' OR 
          va.Produit = '{max_margin_product['produit']}' OR 
          vp.Produit = '{max_margin_product['produit']}'
    """

    df_distribution = execute_sql_query(conn, query_margin_distribution)
    
    # Calcul des totaux pour chaque pays
    total_marges = df_distribution[['marge_france', 'marge_allemagne', 'marge_pologne']].astype(float).values.flatten()
    labels = ['France', 'Allemagne', 'Pologne']

    # Appeler la fonction pour générer le graphique
    ax = axs[1, 1] 
    generate_margin_distribution_pie_chart(ax, total_marges, max_margin_product['lib_produit'], labels)


def main():
    # Chargement des données depuis Excel
    excel_file = "Test_Données.xlsx"
    sheets = load_data_from_excel(excel_file)
    conn = load_data_to_postgresql(sheets)

    # Création de la figure avec plusieurs sous-graphes
    fig, axs = plt.subplots(2, 2, figsize=(13, 9))  # 2 lignes, 2 colonnes

    # Section 3
    section_3(axs, conn)

    # # Section 4 - Marge par produit et répartition de la marge par pays
    section_4(axs, conn)

    plt.tight_layout()
    plt.subplots_adjust(hspace=0.7) 
    plt.savefig("output.png")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")
    except KeyboardInterrupt:
        print("Interruption de l'utilisateur.")
