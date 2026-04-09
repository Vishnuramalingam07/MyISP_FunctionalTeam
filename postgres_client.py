"""
postgres_client.py
──────────────────
PostgreSQL client that mimics the Supabase interface for compatibility.

Usage anywhere in the project:
    from postgres_client import postgres as database
    data = database.table('attendance_records').select('*').execute()
"""

import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import SimpleConnectionPool
from typing import List, Dict, Any, Optional
import os
from dotenv import load_dotenv

# Load .env if present
load_dotenv()

# Database connection parameters
DB_HOST = os.environ.get("POSTGRES_HOST", "localhost")
DB_PORT = os.environ.get("POSTGRES_PORT", "5432")
DB_NAME = os.environ.get("POSTGRES_DB", "myisp_tools")
DB_USER = os.environ.get("POSTGRES_USER", "postgres")
DB_PASSWORD = os.environ.get("POSTGRES_PASSWORD", "postgres123")


class QueryResult:
    """Mimics Supabase query result"""
    def __init__(self, data: List[Dict[str, Any]], error: Optional[Exception] = None):
        self.data = data
        self.error = error
        self.count = len(data) if data else 0


class TableQuery:
    """Mimics Supabase table query builder"""
    def __init__(self, pool: SimpleConnectionPool, table_name: str):
        self.pool = pool
        self.table_name = table_name
        self._select_columns = '*'
        self._insert_data = None
        self._upsert_data = None
        self._update_data = None
        self._delete_mode = False
        self._where_conditions = []
        self._order_by = None
        self._limit_value = None
    
    def select(self, columns: str = '*'):
        """Select columns"""
        self._select_columns = columns
        return self
    
    def insert(self, data: Any):
        """Insert data"""
        self._insert_data = data
        return self
    
    def upsert(self, data: Any, on_conflict: str = None):
        """Upsert data (insert or update on conflict)"""
        self._upsert_data = data
        self._on_conflict =on_conflict
        return self
    
    def update(self, data: Dict[str, Any]):
        """Update data"""
        self._update_data = data
        return self
    
    def delete(self):
        """Delete mode"""
        self._delete_mode = True
        return self
    
    def eq(self, column: str, value: Any):
        """Add WHERE column = value condition"""
        self._where_conditions.append((column, '=', value))
        return self
    
    def order(self, column: str, desc: bool = False):
        """Order by column"""
        direction = 'DESC' if desc else 'ASC'
        self._order_by = f"{column} {direction}"
        return self
    
    def limit(self, count: int):
        """Limit results"""
        self._limit_value = count
        return self
    
    def execute(self) -> QueryResult:
        """Execute the query"""
        conn = None
        try:
            conn = self.pool.getconn()
            cursor = conn.cursor(cursor_factory=RealDictCursor)
            
            # Handle different query types
            if self._insert_data is not None:
                return self._execute_insert(cursor, conn)
            elif self._upsert_data is not None:
                return self._execute_upsert(cursor, conn)
            elif self._update_data is not None:
                return self._execute_update(cursor, conn)
            elif self._delete_mode:
                return self._execute_delete(cursor, conn)
            else:
                return self._execute_select(cursor)
                
        except Exception as e:
            if conn:
                conn.rollback()
            return QueryResult([], error=e)
        finally:
            if conn:
                self.pool.putconn(conn)
    
    def _execute_select(self, cursor) -> QueryResult:
        """Execute SELECT query"""
        query = f"SELECT {self._select_columns} FROM {self.table_name}"
        params = []
        
        if self._where_conditions:
            where_clauses = []
            for col, op, val in self._where_conditions:
                where_clauses.append(f"{col} {op} %s")
                params.append(val)
            query += " WHERE " + " AND ".join(where_clauses)
        
        if self._order_by:
            query += f" ORDER BY {self._order_by}"
        
        if self._limit_value:
            query += f" LIMIT {self._limit_value}"
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return QueryResult([dict(row) for row in rows])
    
    def _execute_insert(self, cursor, conn) -> QueryResult:
        """Execute INSERT query"""
        data_list = self._insert_data if isinstance(self._insert_data, list) else [self._insert_data]
        
        for data in data_list:
            columns = ', '.join(data.keys())
            placeholders = ', '.join(['%s'] * len(data))
            values = list(data.values())
            
            query = f"INSERT INTO {self.table_name} ({columns}) VALUES ({placeholders})"
            cursor.execute(query, values)
        
        conn.commit()
        return QueryResult(data_list)
    
    def _execute_upsert(self, cursor, conn) -> QueryResult:
        """Execute UPSERT query (INSERT ... ON CONFLICT)"""
        data_list = self._upsert_data if isinstance(self._upsert_data, list) else [self._upsert_data]
        
        for data in data_list:
            columns = ', '.join(data.keys())
            placeholders = ', '.join(['%s'] * len(data))
            values = list(data.values())
            
            # Determine conflict target (unique constraint columns)
            conflict_target = self._on_conflict if self._on_conflict else self._get_unique_columns()
            
            # Build UPDATE clause for all columns except the conflict target
            update_cols = [k for k in data.keys() if k not in conflict_target.split(',')]
            update_clause = ', '.join([f"{col} = EXCLUDED.{col}" for col in update_cols])
            
            query = f"""
                INSERT INTO {self.table_name} ({columns}) 
                VALUES ({placeholders})
                ON CONFLICT ({conflict_target}) 
                DO UPDATE SET {update_clause}
            """
            cursor.execute(query, values)
        
        conn.commit()
        return QueryResult(data_list)
    
    def _execute_update(self, cursor, conn) -> QueryResult:
        """Execute UPDATE query"""
        set_clause = ', '.join([f"{k} = %s" for k in self._update_data.keys()])
        params = list(self._update_data.values())
        
        query = f"UPDATE {self.table_name} SET {set_clause}"
        
        if self._where_conditions:
            where_clauses = []
            for col, op, val in self._where_conditions:
                where_clauses.append(f"{col} {op} %s")
                params.append(val)
            query += " WHERE " + " AND ".join(where_clauses)
        
        cursor.execute(query, params)
        conn.commit()
        return QueryResult([self._update_data])
    
    def _execute_delete(self, cursor, conn) -> QueryResult:
        """Execute DELETE query"""
        query = f"DELETE FROM {self.table_name}"
        params = []
        
        if self._where_conditions:
            where_clauses = []
            for col, op, val in self._where_conditions:
                where_clauses.append(f"{col} {op} %s")
                params.append(val)
            query += " WHERE " + " AND ".join(where_clauses)
        
        cursor.execute(query, params)
        conn.commit()
        return QueryResult([])
    
    def _get_unique_columns(self) -> str:
        """Get unique constraint columns for this table"""
        # For attendance_records, the unique constraint is (member_name, year, month, day)
        unique_mappings = {
            'attendance_records': 'member_name, year, month, day',
            'authorized_users': 'username',
            'team_members': 'name'
        }
        return unique_mappings.get(self.table_name, 'id')


class PostgresClient:
    """PostgreSQL client that mimics Supabase interface"""
    def __init__(self):
        # Create connection pool
        self.pool = SimpleConnectionPool(
            minconn=1,
            maxconn=10,
            host=DB_HOST,
            port=DB_PORT,
            database=DB_NAME,
            user=DB_USER,
            password=DB_PASSWORD
        )
        
    def table(self, table_name: str) -> TableQuery:
        """Get a query builder for a table"""
        return TableQuery(self.pool, table_name)
    
    def __del__(self):
        """Close all connections when the client is destroyed"""
        if hasattr(self, 'pool'):
            self.pool.closeall()


# Singleton PostgreSQL client
postgres = PostgresClient()
