import pandas as pd
from flask import flash
import os

class CSVLoader:
    def __init__(self, paths):
        self.paths = paths
        self.mapping_949v = {}
        self.mapping_949d = {}
        self.mapping_949x = {}
        self.articles = {}
        self.sonderzeichen = {}
        self.default_values = {}
        self.mapping_905o = {}

    def load_csv_mappings(self):
        try:
            # Mapping 949v
            if os.path.exists(self.paths["csv_mapping_949v"]):
                df = pd.read_csv(self.paths["csv_mapping_949v"])
                if '264$b' in df.columns and '949$v' in df.columns:
                    self.mapping_949v = dict(zip(df['264$b'], df['949$v']))
                    flash("Mapping 949v erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_949v.csv", 'error')

            # Mapping 949d
            if os.path.exists(self.paths["csv_mapping_949d"]):
                df = pd.read_csv(self.paths["csv_mapping_949d"])
                if '905$n' in df.columns and '949$d' in df.columns:
                    self.mapping_949d = dict(zip(df['905$n'], df['949$d']))
                    flash("Mapping 949d erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_949d.csv", 'error')

            # Mapping 949x
            if os.path.exists(self.paths["csv_mapping_949x"]):
                df = pd.read_csv(self.paths["csv_mapping_949x"])
                if '905$n' in df.columns and '949$x' in df.columns:
                    self.mapping_949x = dict(zip(df['905$n'], df['949$x']))
                    flash("Mapping 949x erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_949x.csv", 'error')

            # Mapping 905o
            if os.path.exists(self.paths["csv_mapping_905o"]):
                df = pd.read_csv(self.paths["csv_mapping_905o"])
                if '905$n' in df.columns and '905$o' in df.columns:
                    self.mapping_905o = dict(zip(df['905$n'], df['905$o']))
                    flash("Mapping 905o erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_905o.csv", 'error')

            # Mapping articles
            if os.path.exists(self.paths["csv_mapping_articles"]):
                df = pd.read_csv(self.paths["csv_mapping_articles"])
                if 'article' in df.columns and 'formatted_article' in df.columns:
                    self.articles = dict(zip(df['article'], df['formatted_article']))
                    flash("Mapping articles erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_articles.csv", 'error')

            # Mapping sonderzeichen
            if os.path.exists(self.paths["csv_mapping_sonderzeichen"]):
                df = pd.read_csv(self.paths["csv_mapping_sonderzeichen"])
                if 'original' in df.columns and 'replacement' in df.columns:
                    self.sonderzeichen = dict(zip(df['original'], df['replacement']))
                    flash("Mapping sonderzeichen erfolgreich geladen.", 'success')
                else:
                    flash("Fehlende Spalten in csv_mapping_sonderzeichen.csv", 'error')

            return True

        except Exception as e:
            flash(f"Fehler beim Laden der CSV-Dateien: {str(e)}", 'error')
            print(f"Fehler beim Laden der CSV-Dateien: {str(e)}")
            return False
