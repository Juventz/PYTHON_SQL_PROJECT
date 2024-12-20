import pandas as pd
import sqlite3
import matplotlib.pyplot as plt


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


def load_data_to_sqlite(sheets):
    """Charge les DataFrames dans une base de données SQLite"""
    conn = sqlite3.connect(":memory:")
    for sheet_name, df in sheets.items():
        df.to_sql(sheet_name, conn, if_exists="replace", index=False)
    return conn


def execute_sql_query(conn, query):
    """Exécute une requête SQL et retourne le résultat sous forme de DataFrame"""
    return pd.read_sql_query(query, conn)


def generate_graph_bar_from_dataframe(ax, df, title, x_label, y_label):
    """Génère un graphique à barres à partir d'un DataFrame"""
    ax.bar(df[x_label], df[y_label], color=["skyblue", "green", "red"])
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)


def generate_grouped_bar_chart(ax, df, title, x_label, y_label):
    """Génère un graphique à barres groupées à partir d'un DataFrame"""
    bar_width = 0.35  # Largeur des barres
    index = range(len(df))
    
    # Barres pour 2019 et 2020
    ax.bar(index, df['CA_2019'], width=bar_width, label='2019', color='skyblue')
    ax.bar([i + bar_width for i in index], df['CA_2020'], width=bar_width, label='2020', color='orange')
    
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.set_xticks([i + bar_width / 2 for i in index])
    ax.set_xticklabels(df['Pays'])
    ax.legend()
    
    # Calculer l'augmentation
    df['Increase'] = df['CA_2020'] - df['CA_2019']
    max_increase_country = df.loc[df['Increase'].idxmax()]
    
    # Calculer le pourcentage d'augmentation
    increase_value = max_increase_country['Increase']
    increase_percentage = (increase_value / max_increase_country['CA_2019']) * 100

    # Ajouter une annotation pour le pays avec la plus forte augmentation
    ax.text(0.5, -0.3, f"Le pays avec la plus forte augmentation de CA est : {max_increase_country['Pays']} - {increase_value} ({increase_percentage:.2f}%)",
            ha='center', va='center', transform=ax.transAxes, fontsize=11)


def section_3(axs, conn):
    """Génère les graphiques pour la section 3"""
    # Graphique 1 : Chiffre d'affaires par pays
    query_ca_by_country = """
    SELECT Pays, SUM(CA) AS Total_CA
    FROM (
        SELECT 'France' AS Pays, Ca FROM Vente_France
        UNION ALL
        SELECT 'Allemagne' AS Pays, Ca FROM Vente_Allemagne
        UNION ALL
        SELECT 'Pologne' AS Pays, Ca FROM Vente_Pologne
    )
    GROUP BY Pays
    ORDER BY Total_CA DESC
    """
    df_ca = execute_sql_query(conn, query_ca_by_country)
    generate_graph_bar_from_dataframe(axs[0, 0], df_ca, "Chiffre d'affaires par Pays", "Pays", "Total_CA")

    # Ajouter une annotation pour le pays avec le CA le plus élevé
    max_ca_country = df_ca.iloc[0]
    max_ca_value = max_ca_country["Total_CA"]
    total_ca_sum = df_ca["Total_CA"].sum()
    ca_percentage = (max_ca_value / total_ca_sum) * 100
    axs[0, 0].text(0.5, -0.3, f"Le pays avec le CA le plus élevé est : {max_ca_country['Pays']} - {max_ca_value} ({ca_percentage:.2f}%)",
                   ha='center', va='center', transform=axs[0, 0].transAxes, fontsize=11)

    # Graphique 2 : Évolution du CA
    query_ca_evolution = """
    SELECT Pays,
              SUM(CASE WHEN Annee = 2020 THEN CA ELSE 0 END) AS CA_2020,
              SUM(CASE WHEN Annee = 2019 THEN CA ELSE 0 END) AS CA_2019
    FROM (
        SELECT 'France' AS Pays, Annee, Ca FROM Vente_France
        UNION ALL
        SELECT 'Allemagne' AS Pays, Annee, Ca FROM Vente_Allemagne
        UNION ALL
        SELECT 'Pologne' AS Pays, Annee, Ca FROM Vente_Pologne
    )
    GROUP BY Pays
    ORDER BY Pays
    """
    df_evolution = execute_sql_query(conn, query_ca_evolution)
    generate_grouped_bar_chart(axs[0, 1], df_evolution, "Évolution du Chiffre d'Affaires", "Pays", "CA")


def generate_margin_bar_chart(ax, df_marge):
    """Génère un graphique à barres de la marge par produit."""
    ax.bar(df_marge['Lib_Produit'], df_marge['Marge_Totale'], color='red')
    ax.set_title("Marge par produit (Tous les pays)")
    ax.set_xlabel("Produit")
    ax.set_ylabel("Marge")
    ax.tick_params(axis='x', rotation=45)  # Pour mieux afficher les noms des produits

    # Trouver le produit ayant généré le plus de marge
    max_margin_product = df_marge.iloc[0]
    max_margin_value = max_margin_product['Marge_Totale']
    total_margin = df_marge['Marge_Totale'].sum()
    margin_percentage = (max_margin_value / total_margin) * 100

    # Ajouter une annotation sous le graphique avec un décalage
    ax.text(0.5, -0.3, 
            f"Le produit avec la marge la plus élevée est : {max_margin_product['Lib_Produit']} - {max_margin_value} ({margin_percentage:.2f}%)",
            ha='center', va='center', transform=ax.transAxes, fontsize=10)


def generate_margin_distribution_pie_chart(ax, df_distribution, product_name):
    """Génère un graphique en camembert de la répartition de la marge par pays."""
    values = df_distribution[['Marge_France', 'Marge_Allemagne', 'Marge_Pologne']].values.flatten()
    labels = df_distribution['Pays'].values

    # Tracer le graphique en camembert
    ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
    ax.set_title(f"Répartition de la marge pour : {product_name}")

    # Trouver le pays avec la plus grande part de marge
    max_margin_country = df_distribution.loc[df_distribution[['Marge_France', 'Marge_Allemagne', 'Marge_Pologne']].values.flatten().argmax()]
    max_country_margin = max(max_margin_country[['Marge_France', 'Marge_Allemagne', 'Marge_Pologne']])
    total_margin = max_margin_country.sum()  # Utiliser la somme des marges
    margin_percentage = (max_country_margin / total_margin) * 100

    # Ajouter une annotation en dessous du graphique
    ax.text(0, -1.2, f"Le pays avec la plus grande part de marge : {max_margin_country['Pays']} - {max_country_margin} ({margin_percentage:.2f}%)",
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
    FROM Referentiel_Produit rp
    LEFT JOIN (
        SELECT Produit, CA FROM Vente_France
    ) vf ON rp.Produit = vf.Produit
    LEFT JOIN (
        SELECT Produit, CA FROM Vente_Allemagne
    ) va ON rp.Produit = va.Produit
    LEFT JOIN (
        SELECT Produit, CA FROM Vente_Pologne
    ) vp ON rp.Produit = vp.Produit
    JOIN Cout c ON rp.Produit = c.Produit
    GROUP BY rp.Produit, rp.Lib_Produit
    ORDER BY Marge_Totale DESC
    """

    df_marge = execute_sql_query(conn, query_cost_per_product)

    # Graphique de la marge par produit
    generate_margin_bar_chart(axs[1, 0], df_marge)

    # Maintenant, créer le graphique de répartition de la marge par pays
    max_margin_product = df_marge.iloc[0]
    query_margin_distribution = f"""
    SELECT Pays,
           SUM(CA - Cout_France) AS Marge_France,
           SUM(CA - Cout_Allemagne) AS Marge_Allemagne,
           SUM(CA - Cout_Pologne) AS Marge_Pologne
    FROM (
        SELECT 'France' AS Pays, Produit, CA FROM Vente_France WHERE Produit = {max_margin_product['Produit']}
        UNION ALL
        SELECT 'Allemagne' AS Pays, Produit, CA FROM Vente_Allemagne WHERE Produit = {max_margin_product['Produit']}
        UNION ALL
        SELECT 'Pologne' AS Pays, Produit, CA FROM Vente_Pologne WHERE Produit = {max_margin_product['Produit']}
    ) v
    JOIN Cout c ON v.Produit = c.Produit
    GROUP BY Pays
    """

    df_distribution = execute_sql_query(conn, query_margin_distribution)
    print("debug", df_distribution)

    # Vérification des valeurs
    if df_distribution.empty or df_distribution[['Marge_France', 'Marge_Allemagne', 'Marge_Pologne']].isnull().any().any():
        print("Erreur : Les données de distribution de marge sont vides ou contiennent des valeurs nulles.")
        return

    # Graphique en camembert pour la répartition de la marge par pays
    generate_margin_distribution_pie_chart(axs[1, 1], df_distribution, max_margin_product['Lib_Produit'])
 

def main():
    # Chargement des données depuis Excel
    excel_file = "QuadraticData_Test_Données.xlsx"
    sheets = load_data_from_excel(excel_file)
    conn = load_data_to_sqlite(sheets)

    # Création de la figure avec plusieurs sous-graphes
    fig, axs = plt.subplots(2, 2, figsize=(13, 9))  # 2 lignes, 2 colonnes

    # Section 3
    section_3(axs, conn)

    # Section 4 - Marge par produit et répartition de la marge par pays
    section_4(axs, conn)

    plt.tight_layout()
    plt.subplots_adjust(hspace=0.7)  # Ajuster l'espace entre les lignes
    plt.show()


if __name__ == "__main__":
    main()