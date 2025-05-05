import pymysql
from bcrypt import hashpw, gensalt, checkpw

class Auth:
    def __init__(self, db_connection):
        """
        初始化身份验证模块。
        :param db_connection: 数据库连接对象
        """
        self.conn = db_connection
        self.cursor = self.conn.cursor()

    def register(self, username, password, role="user"):
        """
        注册新用户。
        :param username: 用户名
        :param password: 密码
        :param role: 用户角色，默认为 "user"
        :return: 注册结果
        """
        try:
            hashed_password = hashpw(password.encode('utf-8'), gensalt())
            query = "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)"
            self.cursor.execute(query, (username, hashed_password, role))
            self.conn.commit()
            return True, "用户注册成功！"
        except pymysql.IntegrityError:
            return False, "用户名已存在！"
        except Exception as e:
            return False, f"注册失败：{e}"

    def login(self, username, password):
        """
        用户登录。
        :param username: 用户名
        :param password: 密码
        :return: 登录结果和角色
        """
        try:
            query = "SELECT password, role FROM users WHERE username = %s"
            self.cursor.execute(query, (username,))
            result = self.cursor.fetchone()
            if result:
                stored_password, role = result
                if checkpw(password.encode('utf-8'), stored_password.encode('utf-8')):
                    return True, role
            return False, "用户名或密码错误！"
        except Exception as e:
            return False, f"登录失败：{e}"

    def verify_role(self, username, required_role):
        """
        验证用户角色。
        :param username: 用户名
        :param required_role: 所需角色
        :return: 是否具有权限
        """
        try:
            query = "SELECT role FROM users WHERE username = %s"
            self.cursor.execute(query, (username,))
            result = self.cursor.fetchone()
            if result and result[0] == required_role:
                return True
            return False
        except Exception as e:
            return False

    def __del__(self):
        """
        关闭数据库连接。
        """
        self.conn.close()
