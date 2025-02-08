import yaml

class settings:
    def __init__(self, **entries):
        self.__dict__.update(entries)

class configuration:
    def cargar_configuracion(self):
        with open('config.yaml', 'r') as archivo: #C:\\nationalsoft\\FacturaPronta\\
            configuracion = yaml.safe_load(archivo)
            return settings(**configuracion)