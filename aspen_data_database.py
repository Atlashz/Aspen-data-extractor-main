#!/usr/bin/env python3
"""
Aspen Data Database Manager

用于存储和管理从 Aspen Plus 提取的工艺数据的数据库系统。
支持每次运行时覆盖之前的数据，为后续的 TEA 计算提供快速数据访问。

Features:
- SQLite 数据库存储
- 结构化数据模型
- 数据版本控制
- 快速查询接口
- 完整的备份和恢复功能
- Enhanced I-N column support for heat exchangers

Author: TEA Analysis Framework
Date: 2025-07-26
Version: 1.1 - Enhanced I-N Column Support
"""

import sqlite3
import json
import os
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, asdict
import pandas as pd

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')
logger = logging.getLogger(__name__)

@dataclass
class AspenStreamRecord:
    """Aspen 流股数据记录"""
    name: str
    temperature: float
    pressure: float
    mass_flow: float
    volume_flow: float
    molar_flow: float
    composition: str  # JSON string of composition dict
    extraction_time: str

@dataclass
class AspenEquipmentRecord:
    """Aspen 设备数据记录"""
    name: str
    equipment_type: str
    aspen_type: str
    importance: str
    function: str
    parameters: str  # JSON string of parameters dict
    parameter_count: int
    excel_specified: bool
    extraction_time: str

@dataclass
class HeatExchangerRecord:
    """换热器数据记录 - Enhanced with I-N column support"""
    name: str
    duty_kw: float
    area_m2: float
    temperatures: str  # JSON string
    pressures: str     # JSON string
    source: str  # 'excel' or 'aspen'
    extraction_time: str
    # I-N column data
    column_i_data: Optional[float] = None
    column_i_header: Optional[str] = None
    column_j_data: Optional[float] = None
    column_j_header: Optional[str] = None
    column_k_data: Optional[float] = None
    column_k_header: Optional[str] = None
    column_l_data: Optional[float] = None
    column_l_header: Optional[str] = None
    column_m_data: Optional[float] = None
    column_m_header: Optional[str] = None
    column_n_data: Optional[float] = None
    column_n_header: Optional[str] = None
    columns_i_to_n_raw: Optional[str] = None  # JSON string of raw I-N data

@dataclass
class ExtractionSession:
    """数据提取会话记录"""
    session_id: str
    extraction_time: str
    aspen_file_path: str
    hex_file_path: str
    stream_count: int
    equipment_count: int
    hex_count: int
    total_heat_duty_kw: float
    total_heat_area_m2: float
    status: str
    notes: str

class AspenDataDatabase:
    """Aspen 数据数据库管理器 - Enhanced with I-N column support"""
    
    def __init__(self, db_path: str = "aspen_data.db"):
        self.db_path = db_path
        self.connection = None
        self.current_session_id = None
        
        # 初始化数据库
        self._initialize_database()
        
    def _initialize_database(self):
        """初始化数据库表结构"""
        try:
            self.connection = sqlite3.connect(self.db_path)
            self.connection.row_factory = sqlite3.Row  # 允许按列名访问
            
            # 创建表结构
            self._create_tables()
            
            logger.info(f"✅ Database initialized: {self.db_path}")
            
        except Exception as e:
            logger.error(f"Failed to initialize database: {e}")
            raise
    
    def _create_tables(self):
        """创建数据库表"""
        cursor = self.connection.cursor()
        
        # 数据提取会话表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS extraction_sessions (
                session_id TEXT PRIMARY KEY,
                extraction_time TEXT NOT NULL,
                aspen_file_path TEXT,
                hex_file_path TEXT,
                stream_count INTEGER DEFAULT 0,
                equipment_count INTEGER DEFAULT 0,
                hex_count INTEGER DEFAULT 0,
                total_heat_duty_kw REAL DEFAULT 0.0,
                total_heat_area_m2 REAL DEFAULT 0.0,
                status TEXT DEFAULT 'active',
                notes TEXT
            )
        """)
        
        # 流股数据表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS aspen_streams (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                name TEXT NOT NULL,
                temperature REAL,
                pressure REAL,
                mass_flow REAL,
                volume_flow REAL,
                molar_flow REAL,
                composition TEXT,
                extraction_time TEXT,
                stream_category TEXT,
                stream_sub_category TEXT,
                classification_confidence REAL,
                classification_reasoning TEXT,
                custom_name TEXT,
                FOREIGN KEY (session_id) REFERENCES extraction_sessions (session_id)
            )
        """)
        
        # 设备数据表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS aspen_equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                name TEXT NOT NULL,
                equipment_type TEXT,
                aspen_type TEXT,
                importance TEXT,
                function TEXT,
                parameters TEXT,
                parameter_count INTEGER DEFAULT 0,
                excel_specified BOOLEAN DEFAULT 0,
                extraction_time TEXT,
                custom_name TEXT,
                FOREIGN KEY (session_id) REFERENCES extraction_sessions (session_id)
            )
        """)
        
        # 换热器数据表 - Enhanced with hot/cold stream data and I-N column support
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS heat_exchangers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                name TEXT NOT NULL,
                duty_kw REAL DEFAULT 0.0,
                area_m2 REAL DEFAULT 0.0,
                temperatures TEXT,
                pressures TEXT,
                source TEXT DEFAULT 'unknown',
                extraction_time TEXT,
                hot_stream_name TEXT,
                hot_stream_inlet_temp REAL,
                hot_stream_outlet_temp REAL,
                hot_stream_flow_rate REAL,
                hot_stream_composition TEXT,
                cold_stream_name TEXT,
                cold_stream_inlet_temp REAL,
                cold_stream_outlet_temp REAL,
                cold_stream_flow_rate REAL,
                cold_stream_composition TEXT,
                column_i_data REAL,
                column_i_header TEXT,
                column_j_data REAL,
                column_j_header TEXT,
                column_k_data REAL,
                column_k_header TEXT,
                column_l_data REAL,
                column_l_header TEXT,
                column_m_data REAL,
                column_m_header TEXT,
                column_n_data REAL,
                column_n_header TEXT,
                columns_i_to_n_raw TEXT,
                FOREIGN KEY (session_id) REFERENCES extraction_sessions (session_id)
            )
        """)
        
        # 创建索引以提高查询性能
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_streams_name ON aspen_streams(name)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_equipment_name ON aspen_equipment(name)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_equipment_type ON aspen_equipment(equipment_type)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_hex_name ON heat_exchangers(name)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_session_time ON extraction_sessions(extraction_time)")
        
        self.connection.commit()
        logger.info("Database tables created successfully with I-N column support")
    
    def start_new_session(self, aspen_file: str = None, hex_file: str = None) -> str:
        """开始新的数据提取会话"""
        # 生成会话ID
        session_id = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        self.current_session_id = session_id
        
        # 清空现有数据（覆盖模式）
        self._clear_all_data()
        
        # 创建新会话记录
        cursor = self.connection.cursor()
        cursor.execute("""
            INSERT INTO extraction_sessions 
            (session_id, extraction_time, aspen_file_path, hex_file_path, status, notes)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            session_id,
            datetime.now().isoformat(),
            aspen_file,
            hex_file,
            'active',
            'New extraction session started'
        ))
        
        self.connection.commit()
        logger.info(f"✅ New extraction session started: {session_id}")
        
        return session_id
    
    def _clear_all_data(self):
        """清空所有数据表（覆盖模式）"""
        cursor = self.connection.cursor()
        
        # 删除所有数据
        cursor.execute("DELETE FROM heat_exchangers")
        cursor.execute("DELETE FROM aspen_equipment") 
        cursor.execute("DELETE FROM aspen_streams")
        cursor.execute("DELETE FROM extraction_sessions")
        
        self.connection.commit()
        logger.info("🗑️ All existing data cleared")
    
    def store_stream_data(self, streams: Dict[str, Any]):
        """存储流股数据"""
        if not self.current_session_id:
            raise ValueError("No active session. Call start_new_session() first.")
        
        cursor = self.connection.cursor()
        extraction_time = datetime.now().isoformat()
        
        for stream_name, stream_data in streams.items():
            # 处理组成数据
            composition_json = json.dumps(stream_data.get('composition', {}))
            
            cursor.execute("""
                INSERT INTO aspen_streams 
                (session_id, name, temperature, pressure, mass_flow, volume_flow, 
                 molar_flow, composition, extraction_time, stream_category, 
                 stream_sub_category, classification_confidence, custom_name)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                self.current_session_id,
                stream_name,
                stream_data.get('temperature', 0.0),
                stream_data.get('pressure', 0.0),
                stream_data.get('mass_flow', 0.0),
                stream_data.get('volume_flow', 0.0),
                stream_data.get('molar_flow', 0.0),
                composition_json,
                extraction_time,
                stream_data.get('stream_category', None),
                stream_data.get('stream_sub_category', None),
                stream_data.get('classification_confidence', None),
                stream_data.get('custom_name', None)
            ))
        
        self.connection.commit()
        logger.info(f"✅ Stored {len(streams)} stream records")
    
    def store_equipment_data(self, equipment: Dict[str, Any]):
        """存储设备数据"""
        if not self.current_session_id:
            raise ValueError("No active session. Call start_new_session() first.")
        
        cursor = self.connection.cursor()
        extraction_time = datetime.now().isoformat()
        
        for eq_name, eq_data in equipment.items():
            # 处理参数数据
            parameters_json = json.dumps(eq_data.get('parameters', {}))
            
            cursor.execute("""
                INSERT INTO aspen_equipment 
                (session_id, name, equipment_type, aspen_type, importance, 
                 function, parameters, parameter_count, excel_specified, extraction_time, custom_name)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                self.current_session_id,
                eq_name,
                eq_data.get('type', 'Unknown'),
                eq_data.get('aspen_type', 'Unknown'),
                eq_data.get('importance', 'Unknown'),
                eq_data.get('function', 'Unknown'),
                parameters_json,
                eq_data.get('parameter_count', 0),
                eq_data.get('excel_specified', False),
                extraction_time,
                eq_data.get('custom_name', None)
            ))
        
        self.connection.commit()
        logger.info(f"✅ Stored {len(equipment)} equipment records")
    
    def store_hex_data(self, hex_data: Dict[str, Any]):
        """存储换热器数据 - Enhanced with I-N column support"""
        if not self.current_session_id:
            raise ValueError("No active session. Call start_new_session() first.")
        
        cursor = self.connection.cursor()
        extraction_time = datetime.now().isoformat()
        
        heat_exchangers = hex_data.get('heat_exchangers', [])
        
        for hex_info in heat_exchangers:
            cursor.execute("""
                INSERT INTO heat_exchangers 
                (session_id, name, duty_kw, area_m2, temperatures, pressures, source, extraction_time,
                 hot_stream_name, hot_stream_inlet_temp, hot_stream_outlet_temp, hot_stream_flow_rate, hot_stream_composition,
                 cold_stream_name, cold_stream_inlet_temp, cold_stream_outlet_temp, cold_stream_flow_rate, cold_stream_composition,
                 column_i_data, column_i_header, column_j_data, column_j_header, 
                 column_k_data, column_k_header, column_l_data, column_l_header,
                 column_m_data, column_m_header, column_n_data, column_n_header, columns_i_to_n_raw)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                self.current_session_id,
                hex_info.get('name', 'Unknown'),
                hex_info.get('duty', 0.0),
                hex_info.get('area', 0.0),
                json.dumps(hex_info.get('temperatures', {})),
                json.dumps(hex_info.get('pressures', {})),
                'excel',
                extraction_time,
                hex_info.get('hot_stream_name', None),
                hex_info.get('hot_stream_inlet_temp', None),
                hex_info.get('hot_stream_outlet_temp', None),
                hex_info.get('hot_stream_flow_rate', None),
                json.dumps(hex_info.get('hot_stream_composition', {})) if hex_info.get('hot_stream_composition') else None,
                hex_info.get('cold_stream_name', None),
                hex_info.get('cold_stream_inlet_temp', None),
                hex_info.get('cold_stream_outlet_temp', None),
                hex_info.get('cold_stream_flow_rate', None),
                json.dumps(hex_info.get('cold_stream_composition', {})) if hex_info.get('cold_stream_composition') else None,
                # I-N column data
                hex_info.get('column_i_data', None),
                hex_info.get('column_i_header', None),
                hex_info.get('column_j_data', None),
                hex_info.get('column_j_header', None),
                hex_info.get('column_k_data', None),
                hex_info.get('column_k_header', None),
                hex_info.get('column_l_data', None),
                hex_info.get('column_l_header', None),
                hex_info.get('column_m_data', None),
                hex_info.get('column_m_header', None),
                hex_info.get('column_n_data', None),
                hex_info.get('column_n_header', None),
                json.dumps(hex_info.get('columns_i_to_n_raw', {})) if hex_info.get('columns_i_to_n_raw') else None
            ))
        
        self.connection.commit()
        logger.info(f"✅ Stored {len(heat_exchangers)} heat exchanger records with I-N column data")
    
    def finalize_session(self, summary_stats: Dict[str, Any]):
        """完成当前会话并更新统计信息"""
        if not self.current_session_id:
            return
        
        cursor = self.connection.cursor()
        
        cursor.execute("""
            UPDATE extraction_sessions 
            SET stream_count = ?, equipment_count = ?, hex_count = ?,
                total_heat_duty_kw = ?, total_heat_area_m2 = ?, status = ?, notes = ?
            WHERE session_id = ?
        """, (
            summary_stats.get('stream_count', 0),
            summary_stats.get('equipment_count', 0),
            summary_stats.get('hex_count', 0),
            summary_stats.get('total_heat_duty_kw', 0.0),
            summary_stats.get('total_heat_area_m2', 0.0),
            'completed',
            f"Extraction completed successfully at {datetime.now().isoformat()}",
            self.current_session_id
        ))
        
        self.connection.commit()
        logger.info(f"✅ Session finalized: {self.current_session_id}")
    
    def get_latest_session_id(self) -> Optional[str]:
        """获取最新的会话ID"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT session_id FROM extraction_sessions 
            ORDER BY extraction_time DESC LIMIT 1
        """)
        
        result = cursor.fetchone()
        return result['session_id'] if result else None
    
    def get_all_streams(self, session_id: str = None) -> pd.DataFrame:
        """获取所有流股数据"""
        if not session_id:
            session_id = self.get_latest_session_id()
        
        if not session_id:
            logger.warning("No session found")
            return pd.DataFrame()
        
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM aspen_streams WHERE session_id = ?
            ORDER BY name
        """, (session_id,))
        
        rows = cursor.fetchall()
        df = pd.DataFrame([dict(row) for row in rows])
        
        # 解析组成数据
        if not df.empty and 'composition' in df.columns:
            df['composition_dict'] = df['composition'].apply(
                lambda x: json.loads(x) if x else {}
            )
        
        logger.info(f"Retrieved {len(df)} stream records")
        return df
    
    def get_all_equipment(self, session_id: str = None) -> pd.DataFrame:
        """获取所有设备数据"""
        if not session_id:
            session_id = self.get_latest_session_id()
        
        if not session_id:
            logger.warning("No session found")
            return pd.DataFrame()
        
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM aspen_equipment WHERE session_id = ?
            ORDER BY name
        """, (session_id,))
        
        rows = cursor.fetchall()
        df = pd.DataFrame([dict(row) for row in rows])
        
        # 解析参数数据
        if not df.empty and 'parameters' in df.columns:
            df['parameters_dict'] = df['parameters'].apply(
                lambda x: json.loads(x) if x else {}
            )
        
        logger.info(f"Retrieved {len(df)} equipment records")
        return df
    
    def get_all_heat_exchangers(self, session_id: str = None) -> pd.DataFrame:
        """获取所有换热器数据 - Enhanced with I-N column data"""
        if not session_id:
            session_id = self.get_latest_session_id()
        
        if not session_id:
            logger.warning("No session found")
            return pd.DataFrame()
        
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM heat_exchangers WHERE session_id = ?
            ORDER BY name
        """, (session_id,))
        
        rows = cursor.fetchall()
        df = pd.DataFrame([dict(row) for row in rows])
        
        # Parse composition and I-N column data
        if not df.empty:
            if 'hot_stream_composition' in df.columns:
                df['hot_stream_composition_dict'] = df['hot_stream_composition'].apply(
                    lambda x: json.loads(x) if x else {}
                )
            if 'cold_stream_composition' in df.columns:
                df['cold_stream_composition_dict'] = df['cold_stream_composition'].apply(
                    lambda x: json.loads(x) if x else {}
                )
            if 'columns_i_to_n_raw' in df.columns:
                df['columns_i_to_n_raw_dict'] = df['columns_i_to_n_raw'].apply(
                    lambda x: json.loads(x) if x else {}
                )
        
        logger.info(f"Retrieved {len(df)} heat exchanger records with I-N column data")
        return df
    
    def get_i_to_n_column_summary(self, session_id: str = None) -> Dict[str, Any]:
        """获取I-N列数据汇总统计"""
        if not session_id:
            session_id = self.get_latest_session_id()
        
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT 
                COUNT(*) as total_hex,
                COUNT(column_i_data) as column_i_count,
                COUNT(column_j_data) as column_j_count,
                COUNT(column_k_data) as column_k_count,
                COUNT(column_l_data) as column_l_count,
                COUNT(column_m_data) as column_m_count,
                COUNT(column_n_data) as column_n_count,
                AVG(column_i_data) as avg_column_i,
                AVG(column_j_data) as avg_column_j,
                AVG(column_k_data) as avg_column_k,
                AVG(column_l_data) as avg_column_l,
                AVG(column_m_data) as avg_column_m,
                AVG(column_n_data) as avg_column_n
            FROM heat_exchangers 
            WHERE session_id = ?
        """, (session_id,))
        
        result = cursor.fetchone()
        return dict(result) if result else {}
    
    def get_database_summary(self) -> Dict[str, Any]:
        """获取数据库摘要信息"""
        cursor = self.connection.cursor()
        
        # 获取最新会话信息
        cursor.execute("""
            SELECT * FROM extraction_sessions 
            ORDER BY extraction_time DESC LIMIT 1
        """)
        latest_session = cursor.fetchone()
        
        if not latest_session:
            return {"status": "empty", "message": "No data in database"}
        
        session_dict = dict(latest_session)
        
        # 统计各表数据量
        cursor.execute("SELECT COUNT(*) as count FROM aspen_streams WHERE session_id = ?", 
                      (session_dict['session_id'],))
        stream_count = cursor.fetchone()['count']
        
        cursor.execute("SELECT COUNT(*) as count FROM aspen_equipment WHERE session_id = ?", 
                      (session_dict['session_id'],))
        equipment_count = cursor.fetchone()['count']
        
        cursor.execute("SELECT COUNT(*) as count FROM heat_exchangers WHERE session_id = ?", 
                      (session_dict['session_id'],))
        hex_count = cursor.fetchone()['count']
        
        # I-N column data coverage
        i_to_n_summary = self.get_i_to_n_column_summary(session_dict['session_id'])
        
        return {
            "status": "active",
            "latest_session": session_dict,
            "record_counts": {
                "streams": stream_count,
                "equipment": equipment_count,
                "heat_exchangers": hex_count
            },
            "i_to_n_column_coverage": i_to_n_summary,
            "database_path": self.db_path,
            "total_records": stream_count + equipment_count + hex_count
        }
    
    def export_to_json(self, output_file: str = None, session_id: str = None):
        """导出数据库内容到JSON文件"""
        if not output_file:
            output_file = f"excel_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        if not session_id:
            session_id = self.get_latest_session_id()
        
        if not session_id:
            logger.error("No session found for export")
            return False
        
        try:
            # 获取所有数据
            streams_df = self.get_all_streams(session_id)
            equipment_df = self.get_all_equipment(session_id)
            hex_df = self.get_all_heat_exchangers(session_id)
            
            # 转换为字典格式
            export_data = {
                "session_id": session_id,
                "export_time": datetime.now().isoformat(),
                "streams": streams_df.to_dict('records') if not streams_df.empty else [],
                "equipment": equipment_df.to_dict('records') if not equipment_df.empty else [],
                "heat_exchangers": hex_df.to_dict('records') if not hex_df.empty else [],
                "i_to_n_column_summary": self.get_i_to_n_column_summary(session_id),
                "summary": self.get_database_summary()
            }
            
            # 写入JSON文件
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"✅ Database exported to {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"Export failed: {e}")
            return False
    
    def close(self):
        """关闭数据库连接"""
        if self.connection:
            self.connection.close()
            logger.info("Database connection closed")


def test_database_system():
    """测试数据库系统 - Enhanced with I-N column testing"""
    print("\n" + "="*60)
    print("🧪 ASPEN DATA DATABASE SYSTEM TEST - I-N Column Enhanced")
    print("="*60)
    
    # 创建数据库实例
    db = AspenDataDatabase("test_aspen_data.db")
    
    try:
        # 1. 测试会话创建
        print("\n1️⃣ Testing Session Management:")
        session_id = db.start_new_session("test.apw", "BFG-CO2H-HEX.xlsx")
        print(f"   ✅ Session created: {session_id}")
        
        # 2. 测试数据存储
        print("\n2️⃣ Testing Data Storage:")
        
        # 模拟换热器数据 - With I-N column data
        test_hex = {
            "heat_exchangers": [
                {
                    "name": "HEX-001",
                    "duty": 500.0,
                    "area": 125.0,
                    "temperatures": {"hot_in": 200, "hot_out": 150},
                    "pressures": {"shell": 20, "tube": 15},
                    # I-N column data
                    "column_i_data": 200.5,
                    "column_i_header": "Hot Inlet Temp (°C)",
                    "column_j_data": 150.2,
                    "column_j_header": "Hot Outlet Temp (°C)",
                    "column_k_data": 30.0,
                    "column_k_header": "Cold Inlet Temp (°C)",
                    "column_l_data": 80.5,
                    "column_l_header": "Cold Outlet Temp (°C)",
                    "column_m_data": 1000.0,
                    "column_m_header": "Hot Flow Rate (kg/h)",
                    "column_n_data": 1200.0,
                    "column_n_header": "Cold Flow Rate (kg/h)",
                    "columns_i_to_n_raw": {
                        "I": 200.5, "J": 150.2, "K": 30.0, 
                        "L": 80.5, "M": 1000.0, "N": 1200.0
                    }
                }
            ]
        }
        
        db.store_hex_data(test_hex)
        print(f"   ✅ Stored {len(test_hex['heat_exchangers'])} heat exchanger records with I-N data")
        
        # 3. 测试数据检索
        print("\n3️⃣ Testing Data Retrieval:")
        
        hex_df = db.get_all_heat_exchangers()
        print(f"   ✅ Retrieved {len(hex_df)} heat exchanger records")
        
        # 4. 测试I-N列数据汇总
        print("\n4️⃣ Testing I-N Column Summary:")
        i_to_n_summary = db.get_i_to_n_column_summary()
        print(f"   ✅ I-N Column Coverage:")
        for key, value in i_to_n_summary.items():
            print(f"      {key}: {value}")
        
        # 5. 完成会话
        summary_stats = {
            "stream_count": 0,
            "equipment_count": 0,
            "hex_count": len(test_hex['heat_exchangers']),
            "total_heat_duty_kw": 500.0,
            "total_heat_area_m2": 125.0
        }
        
        db.finalize_session(summary_stats)
        print(f"   ✅ Session finalized with I-N column support")
        
        # 6. 测试数据库摘要
        print("\n5️⃣ Database Summary:")
        summary = db.get_database_summary()
        print(f"   📊 Database Status: {summary['status']}")
        print(f"   📁 Database Path: {summary['database_path']}")
        print(f"   🔢 Total Records: {summary['total_records']}")
        print(f"   📋 I-N Column Coverage: {summary.get('i_to_n_column_coverage', {})}")
        
        print("\n" + "="*60)
        print("✅ DATABASE SYSTEM TEST COMPLETED SUCCESSFULLY - I-N Enhanced")
        print("="*60)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Database test failed: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        db.close()


if __name__ == "__main__":
    test_database_system()