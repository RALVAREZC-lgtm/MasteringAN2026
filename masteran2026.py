import sqlite3
import pandas as pd
import hashlib
import json
from datetime import datetime
import streamlit as st

# ==========================================
# 1. BASE DE DATOS Y CARGA DEL PADRÓN ELECTORAL
# ==========================================
class ElectoralDatabase:
    def __init__(self, db_name="ahora_nacion_electoral.db"):
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        
        # Tabla: Locales de Votación (Correlacionado con el Excel Maestro)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS locales_votacion (
            codigo_lv INTEGER PRIMARY KEY,
            odpe TEXT,
            ubigeo INTEGER,
            departamento TEXT,
            provincia TEXT,
            distrito TEXT,
            nombre_local TEXT,
            direccion_local TEXT
        )
        ''')

        # Tabla: Mesas de Sufragio (Incluye orden de CA y características de la mesa)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS mesas_sufragio (
            mesa_sufragio INTEGER PRIMARY KEY,
            codigo_lv INTEGER,
            orden_ca INTEGER,
            numero_ca INTEGER,
            orden_mesas INTEGER,
            aula INTEGER,
            piso INTEGER,
            pabellon INTEGER,
            electores INTEGER,
            electores_discapacidad INTEGER,
            FOREIGN KEY (codigo_lv) REFERENCES locales_votacion(codigo_lv)
        )
        ''')

        # Tabla: Usuarios y Roles Jerárquicos
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            dni TEXT PRIMARY KEY,
            nombre TEXT,
            telefono TEXT,
            rol TEXT, -- 'PERSONERO', 'COORDINADOR_LOCAL', 'COORDINADOR_DISTRITAL', 'COMANDO'
            password_hash TEXT,
            estado TEXT DEFAULT 'ACTIVO'
        )
        ''')

        # Tabla: Asignaciones de Personeros y Check-in GPS
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS asignaciones_personero (
            id_asignacion INTEGER PRIMARY KEY AUTOINCREMENT,
            dni_personero TEXT,
            mesa_sufragio INTEGER,
            checkin_status INTEGER DEFAULT 0,
            checkin_timestamp TEXT,
            latitud REAL,
            longitud REAL,
            FOREIGN KEY (dni_personero) REFERENCES usuarios(dni),
            FOREIGN KEY (mesa_sufragio) REFERENCES mesas_sufragio(mesa_sufragio)
        )
        ''')

        # Tabla: Conteo Rápido y Actas de Segunda Vuelta
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS actas_segunda_vuelta (
            mesa_sufragio INTEGER PRIMARY KEY,
            dni_personero TEXT,
            votos_ahora_nacion INTEGER,
            votos_contendiente INTEGER,
            votos_blancos INTEGER,
            votos_nulos INTEGER,
            votos_impugnados INTEGER,
            total_emitidos INTEGER,
            foto_acta_url TEXT,
            estado_acta TEXT, -- 'VALIDADA', 'OBSERVADA_EXCESO_VOTOS', 'OBSERVADA_SIN_VOTOS'
            fecha_registro TEXT,
            FOREIGN KEY (mesa_sufragio) REFERENCES mesas_sufragio(mesa_sufragio)
        )
        ''')

        # Tabla: Alertas Legal y Pánico
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS incidencias_legales (
            id_incidencia INTEGER PRIMARY KEY AUTOINCREMENT,
            mesa_sufragio INTEGER,
            dni_reporta TEXT,
            tipo_incidencia TEXT,
            descripcion TEXT,
            prioridad TEXT, -- 'ALTA', 'MEDIA', 'BAJA'
            fecha_hora TEXT
        )
        ''')

        self.conn.commit()

    def cargar_excel_maestro(self, excel_path="Maestro_EMR2026.xlsx"):
        """Carga y castea los tipos de datos desde el Excel maestro hacia SQLite"""
        df = pd.read_excel(excel_path, sheet_name='TacnaEMR2026')
        cursor = self.conn.cursor()

        # 1. Cargar Locales de Votación
        locales = df[['CÓDIGO DE LV', 'ODPE', 'UBIGEO', 'DEPARTAMENTO', 'PROVINCIA', 
                      'DISTRITO', 'NOMBRE DE LOCAL DE VOTACIÓN', 'DIRECCION DEL LOCAL DE VOTACIÓN']].drop_duplicates()
        
        for _, r in locales.iterrows():
            cursor.execute('''
            INSERT OR REPLACE INTO locales_votacion VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                int(r['CÓDIGO DE LV']), 
                str(r['ODPE']), 
                int(r['UBIGEO']), 
                str(r['DEPARTAMENTO']), 
                str(r['PROVINCIA']), 
                str(r['DISTRITO']), 
                str(r['NOMBRE DE LOCAL DE VOTACIÓN']), 
                str(r['DIRECCION DEL LOCAL DE VOTACIÓN'])
            ))

        # 2. Cargar Mesas de Sufragio con casting seguro (int nativos de Python)
        mesas = df[['MESA DE SUFRAGIO', 'CÓDIGO DE LV', 'ORDEN DE CA', 'NÚMERO DE CA', 'ORDEN DE MESAS', 
                    'AULA', 'PISO', 'PABELLÓN', 'ELECTORES', 'ELECTORES CON DISCAPACIDAD']]
        
        for _, r in mesas.iterrows():
            cursor.execute('''
            INSERT OR REPLACE INTO mesas_sufragio VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                int(r['MESA DE SUFRAGIO']), 
                int(r['CÓDIGO DE LV']), 
                int(r['ORDEN DE CA']),
                int(r['NÚMERO DE CA']),
                int(r['ORDEN DE MESAS']),
                int(r['AULA']), 
                int(r['PISO']), 
                int(r['PABELLÓN']), 
                int(r['ELECTORES']), 
                int(r['ELECTORES CON DISCAPACIDAD'])
            ))

        self.conn.commit()
        print(f"Padrón Maestro cargado: {len(locales)} locales y {len(mesas)} mesas.")


# ==========================================
# 2. MOTOR DE LOGICA DE NEGOCIO Y SERVICIOS
# ==========================================
class ServicioElectoral:
    def __init__(self, db: ElectoralDatabase):
        self.db = db

    def hash_password(self, password: str) -> str:
        return hashlib.sha256(password.encode()).hexdigest()

    def registrar_usuario(self, dni, nombre, telefono, rol, password):
        cursor = self.db.conn.cursor()
        pwd_hash = self.hash_password(password)
        cursor.execute("INSERT OR REPLACE INTO usuarios VALUES (?, ?, ?, ?, ?, 'ACTIVO')",
                       (str(dni), str(nombre), str(telefono), str(rol), pwd_hash))
        self.db.conn.commit()
        return {"status": "SUCCESS", "message": f"Usuario {nombre} registrado como {rol}."}

    def autenticar_usuario(self, dni, password):
        cursor = self.db.conn.cursor()
        pwd_hash = self.hash_password(password)
        cursor.execute("SELECT dni, nombre, rol FROM usuarios WHERE dni = ? AND password_hash = ?", (str(dni), pwd_hash))
        user = cursor.fetchone()
        if user:
            return {"autenticado": True, "user": {"dni": user[0], "nombre": user[1], "rol": user[2]}}
        return {"autenticado": False, "message": "Credenciales inválidas."}

    def consultar_mesa(self, mesa_sufragio):
        """Módulo: Consulta de Mesas y Ubicación Electoral"""
        cursor = self.db.conn.cursor()
        query = '''
        SELECT m.mesa_sufragio, m.aula, m.piso, m.pabellon, m.electores, m.electores_discapacidad,
               m.orden_ca, m.numero_ca, l.nombre_local, l.direccion_local, l.distrito, l.provincia, l.departamento
        FROM mesas_sufragio m
        JOIN locales_votacion l ON m.codigo_lv = l.codigo_lv
        WHERE m.mesa_sufragio = ?
        '''
        cursor.execute(query, (int(mesa_sufragio),))
        res = cursor.fetchone()
        if res:
            return {
                "mesa": res[0], "aula": res[1], "piso": res[2], "pabellon": res[3],
                "electores": res[4], "discapacidad": res[5],
                "orden_ca": res[6], "numero_ca": res[7],
                "local": res[8], "direccion": res[9], "distrito": res[10],
                "provincia": res[11], "departamento": res[12]
            }
        return None

    def asignar_personero_mesa(self, dni_personero, mesa_sufragio):
        """Módulo: Asignación Territorial de Personeros"""
        cursor = self.db.conn.cursor()
        cursor.execute("INSERT INTO asignaciones_personero (dni_personero, mesa_sufragio) VALUES (?, ?)",
                       (str(dni_personero), int(mesa_sufragio)))
        self.db.conn.commit()
        return {"status": "SUCCESS", "message": f"Personero DNI {dni_personero} asignado a Mesa {mesa_sufragio}."}

    def checkin_personero(self, dni_personero, latitud, longitud):
        """Módulo: Check-in de Asistencia vía GPS"""
        cursor = self.db.conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
        UPDATE asignaciones_personero 
        SET checkin_status = 1, checkin_timestamp = ?, latitud = ?, longitud = ?
        WHERE dni_personero = ?
        ''', (now, float(latitud), float(longitud), str(dni_personero)))
        self.db.conn.commit()
        return {"status": "CHECKIN_OK", "timestamp": now}

    def registrar_acta_segunda_vuelta(self, dni_personero, mesa_sufragio, v_ahora_nacion, v_contendiente, v_blancos, v_nulos, v_impugnados, foto_url):
        """Módulo: Digitalización, Conteo Rápido y Validación Algorítmica de Actas"""
        cursor = self.db.conn.cursor()
        
        cursor.execute("SELECT electores FROM mesas_sufragio WHERE mesa_sufragio = ?", (int(mesa_sufragio),))
        mesa_info = cursor.fetchone()
        if not mesa_info:
            return {"status": "ERROR", "message": "La mesa ingresada no existe en el padrón."}
        
        max_electores = mesa_info[0]
        total_calculado = int(v_ahora_nacion + v_contendiente + v_blancos + v_nulos + v_impugnados)

        # Regla Ágil de Validación
        estado = "VALIDADA"
        if total_calculado > max_electores:
            estado = "OBSERVADA_EXCESO_VOTOS"
        elif total_calculado == 0:
            estado = "OBSERVADA_SIN_VOTOS"

        now = datetime.now().isoformat()
        cursor.execute('''
        INSERT OR REPLACE INTO actas_segunda_vuelta VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (int(mesa_sufragio), str(dni_personero), int(v_ahora_nacion), int(v_contendiente), 
              int(v_blancos), int(v_nulos), int(v_impugnados), total_calculado, str(foto_url), estado, now))
        
        self.db.conn.commit()
        return {
            "status": "REGISTRADO",
            "estado_acta": estado,
            "total_emitidos": total_calculado,
            "max_electores_mesa": max_electores
        }

    def registrar_incidencia(self, mesa_sufragio, dni_reporta, tipo, descripcion, prioridad="ALTA"):
        """Módulo: Botón de Pánico y Registro de Incidencias Legales"""
        cursor = self.db.conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
        INSERT INTO incidencias_legales (mesa_sufragio, dni_reporta, tipo_incidencia, descripcion, prioridad, fecha_hora)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', (int(mesa_sufragio), str(dni_reporta), str(tipo), str(descripcion), str(prioridad), now))
        self.db.conn.commit()
        return {"status": "INCIDENCIA_REGISTRADA", "id": cursor.lastrowid}

    def obtener_resumen_computo_regional(self):
        """Módulo: Dashboard y Tablero de Cómputo Regional en Tiempo Real"""
        cursor = self.db.conn.cursor()
        cursor.execute('''
        SELECT 
            COUNT(mesa_sufragio) as actas_procesadas,
            SUM(votos_ahora_nacion) as total_ahora_nacion,
            SUM(votos_contendiente) as total_contendiente,
            SUM(votos_blancos) as total_blancos,
            SUM(votos_nulos) as total_nulos,
            SUM(total_emitidos) as total_votantes
        FROM actas_segunda_vuelta
        WHERE estado_acta = 'VALIDADA'
        ''')
        res = cursor.fetchone()
        
        cursor.execute("SELECT COUNT(*), SUM(electores) FROM mesas_sufragio")
        totales = cursor.fetchone()
        total_mesas = totales[0]
        total_electores_region = totales[1]

        actas_proc = res[0] if res[0] else 0
        return {
            "total_mesas_region": total_mesas,
            "total_electores_region": total_electores_region,
            "actas_procesadas": actas_proc,
            "porcentaje_escrutado": round((actas_proc / total_mesas) * 100, 2) if total_mesas > 0 else 0,
            "resultados": {
                "AHORA_NACION": res[1] or 0,
                "CONTENDIENTE": res[2] or 0,
                "BLANCOS": res[3] or 0,
                "NULOS": res[4] or 0,
                "TOTAL_EMITIDOS": res[5] or 0
            }
        }


# ==========================================
# 3. GESTIÓN DE INSTANCIA STREAMLIT CACHED
# ==========================================
@st.cache_resource
def inicializar_sistema():
    """Garantiza la instanciación única y la carga protegida del Padrón Maestro"""
    db = ElectoralDatabase()
    cursor = db.conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM mesas_sufragio")
    if cursor.fetchone()[0] == 0:
        db.cargar_excel_maestro("Maestro_EMR2026.xlsx")
    return db


# ==========================================
# 4. PRUEBA DE FUNCIONALIDAD
# ==========================================
if __name__ == "__main__":
    db = inicializar_sistema()
    app = ServicioElectoral(db)

    print("\n--- PRUEBA DE INTEGRIDAD Y SERVICIOS ---")
    # 1. Registrar y Autenticar
    app.registrar_usuario("45879632", "Carlos Mendoza", "952123456", "PERSONERO", "ClaveSegura2026")
    print("Autenticación:", app.autenticar_usuario("45879632", "ClaveSegura2026"))

    # 2. Consultar Mesa 78981 (I.E. Zoila Sabel Cáceres en Tacna)
    mesa_data = app.consultar_mesa(78981)
    print("Consulta Mesa 78981:", json.dumps(mesa_data, indent=2, ensure_ascii=False))

    # 3. Asignar y Check-in
    app.asignar_personero_mesa("45879632", 78981)
    print("Checkin GPS:", app.checkin_personero("45879632", -18.0135, -70.2519))

    # 4. Registrar Acta de Segunda Vuelta
    acta_res = app.registrar_acta_segunda_vuelta(
        dni_personero="45879632",
        mesa_sufragio=78981,
        v_ahora_nacion=165,
        v_contendiente=110,
        v_blancos=8,
        v_nulos=12,
        v_impugnados=0,
        foto_url="https://storage.ahoranacion.pe/actas/78981.jpg"
    )
    print("Registro de Acta:", acta_res)

    # 5. Cómputo General
    print("Resumen Cómputo:", json.dumps(app.obtener_resumen_computo_regional(), indent=2, ensure_ascii=False))