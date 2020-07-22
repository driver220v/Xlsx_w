import psycopg2
import psycopg2.extensions
import data
import re


def connect_db():
    try:
        # Пробное подключениме к базе данных PostgreSQL
        mdb_conn = psycopg2.connect(database='backup_1906',
                                    user='nikolay.yakushev',
                                    host='db.sec.tld',
                                    port='5432',
                                    password='1YkHmeCQJXaK')

        # формируем курсор, с помощью которого можно исполнять SQL-запросы

        # регистрация unicode формата для корректного воcприятия кириллицы
        psycopg2.extensions.register_type(psycopg2.extensions.UNICODE, mdb_conn)
        # формат сессии
        # mdb_conn.set_session(readonly=True, autocommit=False)
        # запись в существующий словарь для дальнейшего вызова из других модулей программы
        data.work_dicts['curs_db'] = mdb_conn
        return True

    except Exception as e:
        print('exception=', e)
        return False


def qdb(query, type_status=None, debug_status=None):
    result = None
    # формируем курсор, с помощью которого можно исполнять SQL-запросы
    cursor_db = data.work_dicts['curs_db'].cursor()
    if type_status == 1:
        result = cursor_db.fetchone()
    elif type_status == 3:
        print('qdb_3')
        if debug_status is True:
            cursor_db.execute(query)
            return 1
        cursor_db.execute(query)
        data.work_dicts['curs_db'].commit()
        result = 1
    else:
        cursor_db.execute(query)
        result = cursor_db.fetchall()

    if result is not None:
        return result
    else:
        return []


def clear_sql(sql_query):
    sql = re.sub("'Null'", 'Null', sql_query)
    return sql
