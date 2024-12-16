import os

class PathManager:
    def __init__(self):
        self.project_dir = os.path.dirname(os.path.abspath(__file__))

    def get_paths(self):
        paths = {
            "output_file": os.path.join(self.project_dir, 'output' 'output.xlsx'),
            "input_dir": os.path.join(self.project_dir, 'uploads'),
            "csv_mapping_949v": os.path.join(self.project_dir, 'Mapping', 'mapping_949v.csv'),
            "csv_mapping_articles": os.path.join(self.project_dir, 'Mapping', 'mapping_articles.csv'),
            "csv_mapping_sonderzeichen": os.path.join(self.project_dir, 'Mapping', 'mapping_sonderzeichen.csv'),
            "csv_mapping_949d": os.path.join(self.project_dir, 'Mapping', 'mapping_949d.csv'),
            "csv_mapping_949x": os.path.join(self.project_dir, 'Mapping', 'mapping_949x.csv'),
            "csv_mapping_905o": os.path.join(self.project_dir, 'Mapping', 'mapping_905o.csv'),
        }
        return paths