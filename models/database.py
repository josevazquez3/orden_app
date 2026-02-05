"""
Modelo de Base de Datos
Sistema de Órdenes del Día - Colegio de Médicos
"""

import sqlite3
from typing import List, Dict, Optional


class Database:
    """Maneja todas las operaciones de base de datos"""
    
    def __init__(self, db_path: str = "orden_dia.db"):
        self.db_path = db_path
        self.crear_tablas()
    
    def get_connection(self):
        """Obtiene conexión a la base de datos"""
        return sqlite3.connect(self.db_path)
    
    def crear_tablas(self):
        """Crea las tablas necesarias"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Tabla de temas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS temas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                descripcion TEXT NOT NULL,
                categoria TEXT,
                activo INTEGER DEFAULT 1,
                fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Tabla de reuniones
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS reuniones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                hora TEXT,
                lugar TEXT,
                sede TEXT,
                tipo TEXT CHECK(tipo IN ('presencial', 'virtual')),
                fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Tabla de delegados
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS delegados (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT,
                nombre TEXT NOT NULL,
                apellido TEXT NOT NULL,
                distrito TEXT,
                titular INTEGER DEFAULT 1,
                activo INTEGER DEFAULT 1
            )
        """)
        
        # Tabla orden_dia
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS orden_dia (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reunion_id INTEGER NOT NULL,
                tema_id INTEGER NOT NULL,
                numero_orden INTEGER NOT NULL,
                observaciones TEXT,
                FOREIGN KEY (reunion_id) REFERENCES reuniones(id),
                FOREIGN KEY (tema_id) REFERENCES temas(id)
            )
        """)
        
        # Tabla firmas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS firmas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reunion_id INTEGER NOT NULL,
                cargo TEXT NOT NULL,
                delegado_id INTEGER NOT NULL,
                FOREIGN KEY (reunion_id) REFERENCES reuniones(id),
                FOREIGN KEY (delegado_id) REFERENCES delegados(id)
            )
        """)
        
        conn.commit()
        conn.close()
        print("[OK] Tablas creadas correctamente")
    
    # === MÉTODOS PARA TEMAS ===
    
    def agregar_tema(self, descripcion: str, categoria: str = "") -> int:
        """Agrega un nuevo tema"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO temas (descripcion, categoria) VALUES (?, ?)",
            (descripcion, categoria)
        )
        tema_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return tema_id
    
    def obtener_temas(self, solo_activos: bool = True) -> List[Dict]:
        """Obtiene lista de temas"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = "SELECT id, descripcion, categoria, activo FROM temas"
        if solo_activos:
            query += " WHERE activo = 1"
        query += " ORDER BY descripcion"
        
        cursor.execute(query)
        temas = []
        for row in cursor.fetchall():
            temas.append({
                'id': row[0],
                'descripcion': row[1],
                'categoria': row[2],
                'activo': row[3]
            })
        
        conn.close()
        return temas
    
    def obtener_tema(self, tema_id: int) -> Optional[Dict]:
        """Obtiene un tema por ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, descripcion, categoria, activo FROM temas WHERE id = ?",
            (tema_id,)
        )
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {
                'id': row[0],
                'descripcion': row[1],
                'categoria': row[2],
                'activo': row[3]
            }
        return None
    
    def modificar_tema(self, tema_id: int, descripcion: str, categoria: str = "") -> bool:
        """Modifica un tema"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE temas SET descripcion = ?, categoria = ? WHERE id = ?",
            (descripcion, categoria, tema_id)
        )
        affected = cursor.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    
    def eliminar_tema(self, tema_id: int) -> bool:
        """Elimina (desactiva) un tema"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE temas SET activo = 0 WHERE id = ?", (tema_id,))
        affected = cursor.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    
    def obtener_historial_tema(self, tema_id: int) -> List[Dict]:
        """Obtiene el historial de un tema"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                r.fecha,
                r.lugar,
                r.sede,
                r.tipo,
                od.numero_orden
            FROM orden_dia od
            JOIN reuniones r ON od.reunion_id = r.id
            WHERE od.tema_id = ?
            ORDER BY r.fecha DESC
        """, (tema_id,))
        
        historial = []
        for row in cursor.fetchall():
            historial.append({
                'fecha': row[0],
                'lugar': row[1],
                'sede': row[2],
                'tipo': row[3],
                'numero_orden': row[4]
            })
        
        conn.close()
        return historial
    
    def obtener_estadisticas_tema(self, tema_id: int) -> Dict:
        """Obtiene estadísticas de un tema"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT COUNT(*) FROM orden_dia WHERE tema_id = ?
        """, (tema_id,))
        cantidad_usos = cursor.fetchone()[0]
        
        cursor.execute("""
            SELECT MIN(r.fecha), MAX(r.fecha)
            FROM orden_dia od
            JOIN reuniones r ON od.reunion_id = r.id
            WHERE od.tema_id = ?
        """, (tema_id,))
        fechas = cursor.fetchone()
        
        conn.close()
        
        return {
            'cantidad_usos': cantidad_usos,
            'primera_fecha': fechas[0] if fechas[0] else None,
            'ultima_fecha': fechas[1] if fechas[1] else None
        }
    
    # === MÉTODOS PARA DELEGADOS ===
    
    def agregar_delegado(self, titulo: str, nombre: str, apellido: str, 
                        distrito: str, titular: bool = True) -> int:
        """Agrega un nuevo delegado"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO delegados (titulo, nombre, apellido, distrito, titular) VALUES (?, ?, ?, ?, ?)",
            (titulo, nombre, apellido, distrito, 1 if titular else 0)
        )
        delegado_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return delegado_id
    
    def obtener_delegados(self, solo_activos: bool = True, solo_titulares: bool = False) -> List[Dict]:
        """Obtiene lista de delegados"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = "SELECT id, titulo, nombre, apellido, distrito, titular, activo FROM delegados WHERE 1=1"
        if solo_activos:
            query += " AND activo = 1"
        if solo_titulares:
            query += " AND titular = 1"
        query += " ORDER BY id"  # Ordenar por ID para mantener orden consistente
        
        cursor.execute(query)
        delegados = []
        for row in cursor.fetchall():
            delegados.append({
                'id': row[0],
                'titulo': row[1],
                'nombre': row[2],
                'apellido': row[3],
                'distrito': row[4],
                'titular': row[5],
                'activo': row[6]
            })
        
        conn.close()
        return delegados
    
    def modificar_delegado(self, delegado_id: int, titulo: str, nombre: str, 
                          apellido: str, distrito: str, titular: bool) -> bool:
        """Modifica un delegado"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE delegados SET titulo = ?, nombre = ?, apellido = ?, distrito = ?, titular = ? WHERE id = ?",
            (titulo, nombre, apellido, distrito, 1 if titular else 0, delegado_id)
        )
        affected = cursor.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    
    def eliminar_delegado(self, delegado_id: int) -> bool:
        """Elimina (desactiva) un delegado"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE delegados SET activo = 0 WHERE id = ?", (delegado_id,))
        affected = cursor.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    
    # === MÉTODOS PARA REUNIONES ===
    
    def agregar_reunion(self, fecha: str, hora: str, lugar: str, sede: str, tipo: str) -> int:
        """Agrega una nueva reunión"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO reuniones (fecha, hora, lugar, sede, tipo) VALUES (?, ?, ?, ?, ?)",
            (fecha, hora, lugar, sede, tipo)
        )
        reunion_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return reunion_id
    
    def agregar_tema_orden_dia(self, reunion_id: int, tema_id: int, numero_orden: int) -> int:
        """Agrega un tema al orden del día"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO orden_dia (reunion_id, tema_id, numero_orden) VALUES (?, ?, ?)",
            (reunion_id, tema_id, numero_orden)
        )
        orden_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return orden_id
    
    def guardar_firmas(self, reunion_id: int, presidente_id: int, secretario_id: int) -> bool:
        """Guarda las firmas de una reunión"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute("DELETE FROM firmas WHERE reunion_id = ?", (reunion_id,))
            
            cursor.execute(
                "INSERT INTO firmas (reunion_id, cargo, delegado_id) VALUES (?, ?, ?)",
                (reunion_id, 'Presidente', presidente_id)
            )
            cursor.execute(
                "INSERT INTO firmas (reunion_id, cargo, delegado_id) VALUES (?, ?, ?)",
                (reunion_id, 'Secretario General', secretario_id)
            )
            
            conn.commit()
            return True
        except Exception as e:
            conn.rollback()
            print(f"Error guardando firmas: {e}")
            return False
        finally:
            conn.close()
    
    def cargar_datos_iniciales(self):
        """Carga los 10 delegados iniciales del PDF"""
        if len(self.obtener_delegados()) > 0:
            return
        
        delegados_iniciales = [
            ("Dr.", "JULIO C.", "MORENO", "Dist. I"),
            ("Dr.", "JORGE E.", "AGUGLIARO", "Dist. II"),
            ("Dr.", "MAURICIO", "ESKINAZI", "Dist. III"),
            ("Dr.", "RUBEN H.", "TUCCI", "Dist. IV"),
            ("Dr.", "JULIO D.", "DUNOGENT", "Dist. V"),
            ("Dr.", "JORGE OSCAR", "LUSARDI", "Dist. VI"),
            ("Dr.", "HORACIO MARIO", "CARDUS", "Dist. VII"),
            ("Dr.", "TOMAS", "GUANELLA", "Dist. VIII"),
            ("Dr.", "GUSTAVO", "ARTURI", "Dist. IX"),
            ("Dra.", "ROSA ANA", "DE FINO", "Dist. X"),
        ]
        
        for titulo, nombre, apellido, distrito in delegados_iniciales:
            self.agregar_delegado(titulo, nombre, apellido, distrito, True)
        
        print("[OK] Delegados iniciales cargados")
    
    # === MÉTODOS PARA REUNIONES ===
    
    def obtener_reuniones(self) -> List[Dict]:
        """Obtiene todas las reuniones del historial"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                r.id,
                r.fecha,
                r.hora,
                r.lugar,
                r.tipo,
                COUNT(od.id) as cantidad_temas
            FROM reuniones r
            LEFT JOIN orden_dia od ON r.id = od.reunion_id
            GROUP BY r.id
            ORDER BY r.fecha DESC
        """)
        
        reuniones = []
        for row in cursor.fetchall():
            reuniones.append({
                'id': row[0],
                'fecha': row[1],
                'hora': row[2],
                'lugar': row[3],
                'tipo': row[4],
                'cantidad_temas': row[5]
            })
        
        conn.close()
        return reuniones
    
    def buscar_reuniones(self, termino_busqueda: str) -> List[Dict]:
        """Busca reuniones por tema o fecha"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Buscar por fecha (YYYY-MM-DD o DD/MM/YYYY)
        # o por descripción de tema
        cursor.execute("""
            SELECT DISTINCT
                r.id,
                r.fecha,
                r.hora,
                r.lugar,
                r.tipo,
                COUNT(od.id) as cantidad_temas
            FROM reuniones r
            LEFT JOIN orden_dia od ON r.id = od.reunion_id
            LEFT JOIN temas t ON od.tema_id = t.id
            WHERE r.fecha LIKE ? OR t.descripcion LIKE ?
            GROUP BY r.id
            ORDER BY r.fecha DESC
        """, (f"%{termino_busqueda}%", f"%{termino_busqueda}%"))
        
        reuniones = []
        for row in cursor.fetchall():
            reuniones.append({
                'id': row[0],
                'fecha': row[1],
                'hora': row[2],
                'lugar': row[3],
                'tipo': row[4],
                'cantidad_temas': row[5]
            })
        
        conn.close()
        return reuniones
    
    def obtener_temas_reunion(self, reunion_id: int) -> List[Dict]:
        """Obtiene todos los temas de una reunión específica"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                t.id,
                t.descripcion,
                t.categoria,
                od.numero_orden
            FROM orden_dia od
            JOIN temas t ON od.tema_id = t.id
            WHERE od.reunion_id = ?
            ORDER BY od.numero_orden
        """, (reunion_id,))
        
        temas = []
        for row in cursor.fetchall():
            temas.append({
                'id': row[0],
                'descripcion': row[1],
                'categoria': row[2],
                'numero_orden': row[3]
            })
        
        conn.close()
        return temas
    
    def eliminar_reunion(self, reunion_id: int) -> bool:
        """Elimina una reunión y su orden del día"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            # Primero eliminar el orden del día
            cursor.execute("DELETE FROM orden_dia WHERE reunion_id = ?", (reunion_id,))
            
            # Luego eliminar las firmas
            cursor.execute("DELETE FROM firmas WHERE reunion_id = ?", (reunion_id,))
            
            # Finalmente eliminar la reunión
            cursor.execute("DELETE FROM reuniones WHERE id = ?", (reunion_id,))
            
            affected = cursor.rowcount
            conn.commit()
            return affected > 0
        except Exception as e:
            conn.rollback()
            print(f"Error eliminando reunión: {e}")
            return False
        finally:
            conn.close()