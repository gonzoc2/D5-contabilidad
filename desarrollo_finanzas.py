import requests
import re
import zeep
from zeep import Client
from zeep.exceptions import Fault
import zeep.exceptions
from zeep.transports import Transport
from requests.auth import HTTPBasicAuth
import zeep.xsd
import pandas as pd
import io

class ReportResponse:
    @property
    def reportBytes() -> bytes: ...
    @property
    def reportContentType() -> str: ...

class Sesion:
    def __init__(self, user: str, password: str, server: str):
        self.user = user
        self.password = password
        self.server = server
        self.session = requests.Session()
        self.session.auth = HTTPBasicAuth(user, password)
        self.client = Client(f"https://{server}.oraclecloud.com:443/xmlpserver/services/ExternalReportWSSService?WSDL", 
                             transport=Transport(session=self.session)
        )
    
    def params(self, params: list[dict]) -> list:
        param_name_value_type = self.client.get_type("ns0:ParamNameValue")
        return [param_name_value_type(**v) for v in params]
    
    def isMissing(self, dynamicParams , params: dict) -> list[str | zeep.xsd.Element]:
        elementName = {element[0] for element in dynamicParams.elements}
        params = {k for k, v in params.items()}
        diff = elementName - params
        return [element[1] for element in dynamicParams.elements if element[0] in diff]
        
    def _integration(self, method: str, keys: dict = None):
        try:
            if not keys: return getattr(self.client.service, method)
            return getattr(self.client.service, method)(**keys)
        except zeep.exceptions.ValidationError as e:
            missingParam = re.findall(r"Missing element (\w+)", str(e))[0]
            for param in keys['reportRequest']['parameterNameValues']['item']:
                if param[missingParam] is None:
                    #print(f'{e} for {param['name']}')
                    param[missingParam] = input(f'Type value for {missingParam}: ')
        return self._integration(method, keys)

    @property
    def validateLogin(self)-> bool:
        try: return self._integration('validateLogin')()
        except Fault: return False
    
    def runReport(self, reportAbsolutePath: str, params: list[dict] = None, *, attributeFormat: str = 'csv', 
                  attributeLocale: str = 'es_MX', attributeCalendar: str = 'gregorian', 
                  attributeTimezone: str = 'America/Mexico_City', attributeUILocale: str = 'es_MX', 
                  byPassCache: bool = True, flattenXML: bool = False, sizeOfDataChunkDownload: int = -1, 
                  appParams: str = '') -> bytes:
        '''Params must be:
           'dataType': [boolean, date, float, integer, string],
            'name': param name,
            'dateFormatString': 'DD-MM-YYYY' id dataType = Date, otherwise no show,
            'values': param value,
            'multiValuesAllowed': Bool,
            'refreshParamOnChange': Bool,
            'selectAll': Bool,
            'templateParam': Bool,
            'useNullForAll': Bool '''
        keys = {
            'appParams': appParams,
            'reportRequest': {
                'reportAbsolutePath': reportAbsolutePath,
                "parameterNameValues": {"item": self.params(params)} if params else None,
                "attributeFormat": attributeFormat,
                "byPassCache": byPassCache,
                "flattenXML": flattenXML,
                "sizeOfDataChunkDownload": sizeOfDataChunkDownload,
                'attributeLocale': attributeLocale,
                'attributeCalendar': attributeCalendar,
                'attributeTimezone': attributeTimezone,
                'attributeUILocale': attributeUILocale,
            }
        }
        return self._integration('runReport', keys)

    def getFolderContent(self, folderAbsolutePath: str):
        args = {'folderAbsolutePath': folderAbsolutePath}
        return self._integration('getFolderContents', args)

    @classmethod
    def login(cls, user: str, pas: str, server: str):
        sesion = cls(user, pas, server)
        if sesion.validateLogin == True:
            return sesion
        return False
        
    def __str__(self):
        return f'Hola {self.user}'
    
def reportes_otm(fecha_ini, fecha_fin):
    USER = 'GCORTES'
    PASS = 'Laku1979#gonza'
    SERVER = 'otmgtm-a621157.otmgtm.us-phoenix-1.ocs'
    CARPETA = '/Custom/ESGARI/Reportes Finanzas/XDO'

    user = Sesion(USER, PASS, SERVER)

    params = [
        {
            'dataType': 'Date',
            'name':'P_FECHA_INI',
            'dateFormatString': 'DD-MM-YYYY',
            'values': fecha_ini.strftime('%m-%d-%Y'),
            'multiValuesAllowed': False,
            'refreshParamOnChange': False,
            'selectAll': False,
            'templateParam': False,
            'useNullForAll': False},
        {
            'dataType': 'Date',
            'name':'P_FECHA_FIN',
            'dateFormatString': 'DD-MM-YYYY',
            'values': fecha_fin.strftime('%m-%d-%Y'),
            'multiValuesAllowed': False,
            'refreshParamOnChange': False,
            'selectAll': False,
            'templateParam': False,
            'useNullForAll': False}
    ]

    contenido = user.getFolderContent(CARPETA)

    try:
        reportes = contenido.item
    except AttributeError:
        reportes = contenido

    dfs = {}
    errores = {}

    for r in reportes:
        if r['type'] != 'Report' or not r['absolutePath'].endswith('.xdo'):
            continue

        nombre = r['displayName']
        ruta = r['absolutePath']

        try:
            salida = user.runReport(ruta, params=params)
            if hasattr(salida, "reportBytes"):
                contenido_csv = salida.reportBytes
            else:
                contenido_csv = salida
            df = pd.read_csv(io.BytesIO(contenido_csv))
            dfs[nombre] = df
        except Exception as e:
            errores[nombre] = str(e)

    # --- Combinar DataFrames por ORDEN_DE_LIBERACION ---

    dataframes_con_orden = [df for df in dfs.values() if "ORDEN_DE_LIBERACION" in df.columns]

    if dataframes_con_orden:

        # Merge progresivo sin renombrar columnas
        df_merged = dataframes_con_orden[0]
        for df in dataframes_con_orden[1:]:
            df_merged = pd.merge(df_merged, df, how="outer", on="ORDEN_DE_LIBERACION")

        # --- Eliminar columnas duplicadas con mismos valores ---
        cols_seen = {}
        cols_to_drop = []

        for col in df_merged.columns:
            if col == "ORDEN_DE_LIBERACION":
                continue
            if col in cols_seen:
                # Compara con la columna ya vista
                if df_merged[col].equals(df_merged[cols_seen[col]]):
                    cols_to_drop.append(col)
            else:
                cols_seen[col] = col  # primera vez que aparece

        df_merged = df_merged.drop(columns=cols_to_drop)


        # Convertir de millas a kilómetros solo donde la unidad de medida es 'MI'
        df_merged.loc[df_merged['UNIDAD_MEDIDA_DISTANCIA_VENTA'] == 'MI', 'DISTANCIA_VENTA'] *= 1.60934


        df_merged.loc[df_merged['TARIFA_VENTA_MONEDA'] == 'USD', 'COSTO_COMPRA_REAL'] *= 20
        df_merged.loc[((df_merged['TARIFA_VENTA_MONEDA'] == 'USD') & (df_merged['COSTO_COMPRA_REAL'] > df_merged['SUBTOTAL_FACTURA'])), 'SUBTOTAL_FACTURA'] *= 20

        # Cambiar la unidad de medida a 'KM' para reflejar el cambio
        df_merged.loc[df_merged['UNIDAD_MEDIDA_DISTANCIA_VENTA'] == 'MI', 'UNIDAD_MEDIDA_DISTANCIA_VENTA'] = 'KM'
        df_merged = df_merged.drop_duplicates(["ORDEN_DE_LIBERACION"])

        return df_merged

if __name__ == '__main__':

    USER = 'GCORTES'
    PASS = 'Laku1979#gonza'
    SERVER = 'otmgtm-a621157.otmgtm.us-phoenix-1.ocs'
    CARPETA = '/Custom/ESGARI/Reportes Finanzas/XDO'
    #folder_path = '/Custom/ESGARI/DEPARTAMENTO FINANZAS/REPORTES QS/XDO/V03/'
    #user = Sesion('rolmedo', 'Mexico.2022', 'ekck.fa.us6')
    user = Sesion(USER, PASS, SERVER)
    #print(user)
    params = [
        {
         'dataType': 'Date',
         'name':'P_FECHA_INI',
         'dateFormatString': 'DD-MM-YYYY',
         'values': '01-01-2025',
         'multiValuesAllowed': False,
         'refreshParamOnChange': False,
         'selectAll': False,
         'templateParam': False,
         'useNullForAll': False},
        {
         'dataType': 'Date',
         'name':'P_FECHA_FIN',
         'dateFormatString': 'DD-MM-YYYY',
         'values': '02-01-2025',
         'multiValuesAllowed': False,
         'refreshParamOnChange': False,
         'selectAll': False,
         'templateParam': False,
         'useNullForAll': False}
    ]
    #h: ReportResponse = user.runReport('/Custom/ESGARI/Qlik/reportesNecesarios/XXRO_EXTRACTOR_GL_REP.xdo', params=params)
    #h = user.runReport("/Custom/ESGARI/AR/XX_AR_FACTURAS_UUID_REP.xdo", params)
    #h: ReportResponse = user.runReport('/Custom/ESGARI/Qlik/Reportes Finanzas/XXRO_EXTRACTOR_GL_REP.xdo')
    #df = pd.read_csv(io.BytesIO(h.reportBytes))
    #df.to_csv('reporte.csv')
    #print(df)
    #print(h)

    # Paso 2: Explorar el contenido de la carpeta
    #folder_path = '/Custom/ESGARI/Reportes Finanzas/XDO'
    print("Login válido:", user.validateLogin)
    print("Ruta consultada:", CARPETA)

    contenido = user.getFolderContent(CARPETA)
    print(contenido)
    #for i, item in enumerate(contenido):
        #print(f"[{i}] Tipo: {type(item)} → Valor: {item}")








