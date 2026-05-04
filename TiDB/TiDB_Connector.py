import mysql.connector

conn = mysql.connector.connect(
    host="gateway01.ap-southeast-1.prod.aws.tidbcloud.com",
    user="2hqkC7CYCyd33Y1.root",
    password="ED6Hesm31WIPzd8e",
    database="test",
    port=4000,
    ssl_ca=r"C:\Sagar\BackTestCode\TiDB\isrgrootx1.pem"
)

print("✅ Connected successfully!")
conn.close()