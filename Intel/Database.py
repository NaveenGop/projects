import pyodbc, re
import WordParse


def main():
    doc = 'jesus.tsv'
    WordParse.main(file_name=doc)
    with open(doc, 'r') as f:
        s = re.split('\n', f.read())
    cnxn = pyodbc.connect(r'DSN=mynewdsn;')
    cursor = cnxn.cursor()
    cursor.execute("""IF OBJECT_ID('tempdb.dbo.##CSV') IS NOT NULL
                        DROP TABLE dbo.##CSV;""")
    cursor.execute(r"""CREATE TABLE ##CSV
                    (
                        Name VARCHAR(50)  NULL 
                        , [ProdCode] VARCHAR(10) NOT NULL
                        , SKU VARCHAR(10) NOT NULL
                        , Density VARCHAR(10) NOT NULL
                        , SKU_Letter VARCHAR(10) NOT NULL
                        , form_factor VARCHAR(200)
                        , PBA VARCHAR(500)  NULL
                        , [Date] DATE  NULL
                        , Link VARCHAR(2000)  NULL
                    );
                    """)
    cursor.commit()
    for x in s:
        y = re.split('\t', x)
        if len(y) < 9:
            continue
        cursor.execute("""insert into ##csv(Name, ProdCode, SKU, Density, SKU_Letter, form_factor, PBA, [Date], Link) values ('%s','%s','%s','%s','%s','%s','%s', CONVERT(DATE, '%s', 102),'%s')""" % (y[0], y[1], y[2], y[3], y[4], y[5], y[6], y[7], y[8]))
        cnxn.commit()
    cursor.execute("""SELECT * FROM ##CSV""")
    row = cursor.fetchone()
    print(2+2)

if __name__ == '__main__':
    main()
