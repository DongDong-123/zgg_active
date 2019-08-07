import pymysql


def connect(parm):
    try:
        conn = pymysql.Connect(host="127.0.0.1", user="root", password="123456", database="zgg_test", port=3306,
                               write_timeout=10, charset="utf8")
        cur = conn.cursor()

        sql = "insert into case_type_price(case_name, price, catlog, state) VALUES(%s,%s,%s,%s)"
        cur.execute(sql, parm)
        conn.commit()

        cur.close()
        conn.close()
        return 1
    except ConnectionError as e:
        print(e)
        return 0


if __name__ == "__main__":
    parm = ("张三", 100.75, 2, 1)
    connect(parm)
