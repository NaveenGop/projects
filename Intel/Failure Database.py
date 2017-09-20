import pyodbc
from datetime import datetime


def main():
    cnxn = pyodbc.connect(r'DSN=mynewdsn')
    cursor = cnxn.cursor()
    connection = pyodbc.connect(r'DSN=mynewdsn')
    slursor = connection.cursor()
    flexor = pyodbc.connect(r'DSN=mynewdsn')
    mursor = flexor.cursor()
    print(datetime.now())
    try:
        cursor.execute(r"""SELECT ProdCode, SKU FROM ##die GROUP BY ##die.ProdCode, ##die.SKU HAVING ##die.Prodcode = 'LJ'""")
    except pyodbc.ProgrammingError:
        cursor.execute(r"""CREATE TABLE tempdb.dbo.##die
                            (
                                ProdCode VARCHAR(4) NOT NULL
                                , SKU VARCHAR(6) NOT NULL
                                , ProcessStep VARCHAR(10) NOT NULL
                                , DieCount INT NOT NULL
                                , SerialCount INT NOT NULL
                            )
                            
                            INSERT INTO ##die(ProdCode, SKU, ProcessStep,DieCount, SerialCount)
                                SELECT
                                    SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber,  3, 2) AS 'ProdCode'
                                    , SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber, 13, 4) AS 'SKU'
                                    , NSGTestDW.dbo.Dim_ProcessStep.ProcessStep
                                    , DiePerTest.DieCount
                                    , COUNT(distinct NSGTestDW.dbo.Fact_EventPVR.SerialNumber)
                                FROM NSGTestDW.dbo.Fact_EventPVR
                                    JOIN (
                                    SELECT
                                        NSGTestDW.dbo.Dim_DiePVR.ComponentEventId
                                        , COUNT(NSGTestDW.dbo.Dim_DiePVR.CE) AS [DieCount]
                                    FROM 
                                        NSGTestDW.dbo.Dim_DiePVR
                                    GROUP BY
                                        NSGTestDW.dbo.Dim_DiePVR.ComponentEventId
                                ) AS DiePerTest
                                        ON DiePerTest.ComponentEventId = NSGTestDW.dbo.Fact_EventPVR.ComponentEventId
                                    JOIN NSGTestDW.dbo.Dim_ProcessStep
                                        ON NSGTestDW.dbo.Fact_EventPVR.ProcessStepId = NSGTestDW.dbo.Dim_ProcessStep.ProcessStepId
                                WHERE
                                     LEN(NSGTestDW.dbo.Fact_EventPVR.SerialNumber) = 18
                                GROUP BY
                                    SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber,  3, 2)
                                    , SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber, 13, 4)
                                    , NSGTestDW.dbo.Dim_ProcessStep.ProcessStep
                                    , DiePerTest.DieCount""")
        cursor.commit()
        cursor.execute(r"""SELECT ProdCode, SKU FROM ##die GROUP BY ##die.ProdCode, ##die.SKU""")
    with open('lenicv2.txt', 'w+b') as f:
        for show in cursor.fetchall():
            curs = cnxn.cursor()
            curs.execute(r"""SELECT
                                ##die.prodcode
                                , ##die.sku
                                , ##die.ProcessStep
                                , ##die.DieCount
                                , ##die.SerialCount
                            FROM 
                                ##die 
                            GROUP BY
                                ##die.prodcode
                                , ##die.sku
                                , ##die.ProcessStep
                                , ##die.DieCount
                                , ##die.SerialCount
                            HAVING
                                ##die.ProdCode = ?
                                AND ##die.SKU = ?
                            ORDER BY 
                                ##die.prodcode
                                , ##die.sku
                                ,##die.SerialCount DESC""", (show[0], show[1]))
            row = curs.fetchone()
            slursor.execute(r"""SELECT TOP 1
                                    Dim_ProcessStep.ProcessStep
                                    , COUNT(distinct Dim_DiePVR.CE) AS [Unique CE]
                                    , COUNT(Dim_DiePVR.CE) AS [Total CE]
                                FROM Fact_EventPVR
                                    INNER JOIN Dim_ProcessStep
                                        ON Fact_EventPVR.ProcessStepId = Dim_ProcessStep.ProcessStepId
                                    INNER JOIN Dim_DiePVR
                                        ON Dim_DiePVR.ComponentEventId = Fact_EventPVR.ComponentEventId
                                WHERE 
                                    SUBSTRING(Fact_EventPVR.SerialNumber, 3, 2) = ?
                                    AND SUBSTRING(Fact_EventPVR.SerialNumber, 13, 4) = ?
                                    AND Dim_ProcessStep.ProcessStep = ?
                                GROUP BY
                                    Fact_EventPVR.ComponentEventId,
                                    Dim_ProcessStep.ProcessStep
                                HAVING
                                    COUNT(Dim_DiePVR.CE) = ?""", (row[0], row[1], row[2], row[3]))
            tow = slursor.fetchall()
            f.write(f"{show[0]} {show[1]} (NORM = {row[3]})\n".encode())
            f.write(b"\tDIE COUNTS BY PROCESS STEP\n")
            f.write(f"\t\tProcessStep = {row[2]} CountCE = {tow[0][2]} NumberOfDrives = {row[4]}\n".encode())
            for row in curs.fetchall():
                slursor.execute(r"""SELECT TOP 1
                                    Dim_ProcessStep.ProcessStep
                                    , COUNT(distinct Dim_DiePVR.CE) AS [Unique CE]
                                    , COUNT(Dim_DiePVR.CE) AS [Total CE]
                                FROM Fact_EventPVR
                                    INNER JOIN Dim_ProcessStep
                                        ON Fact_EventPVR.ProcessStepId = Dim_ProcessStep.ProcessStepId
                                    INNER JOIN Dim_DiePVR
                                        ON Dim_DiePVR.ComponentEventId = Fact_EventPVR.ComponentEventId
                                WHERE 
                                    SUBSTRING(Fact_EventPVR.SerialNumber, 3, 2) = ?
                                    AND SUBSTRING(Fact_EventPVR.SerialNumber, 13, 4) = ?
                                    AND Dim_ProcessStep.ProcessStep = ?
                                GROUP BY
                                    Fact_EventPVR.ComponentEventId,
                                    Dim_ProcessStep.ProcessStep
                                HAVING
                                    COUNT(Dim_DiePVR.CE) = ?""", (row[0], row[1], row[2], row[3]))
                tow = slursor.fetchall()
                f.write(f"\t\tProcessStep = {tow[0][0]} CountCE = {tow[0][2]} NumberOfDrives = {row[4]}\n".encode())
                if tow[0][2] == tow[0][1]*2:
                    f.write(b"\t\t\t-Counted 2 FIDs per CE.\n")
                elif tow[0][2] > tow[0][1]:
                    f.write(b"\t\t\t-Extra CE. They are ")
                    slursor.execute(r"""WITH TBL AS
                                        (
                                            SELECT
                                                Fact_EventPVR.ComponentEventId
                                                , Fact_EventPVR.SerialNumber
                                                , Dim_DiePVR.CE
                                                , COUNT(Dim_DiePVR.CE) AS CountCE
                                            FROM Fact_EventPVR
                                                INNER JOIN Dim_ProcessStep
                                                    ON Fact_EventPVR.ProcessStepId = Dim_ProcessStep.ProcessStepId
                                                INNER JOIN Dim_DiePVR
                                                    ON Dim_DiePVR.ComponentEventId = Fact_EventPVR.ComponentEventId
                                            WHERE
                                                SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber,  3, 2) = ?
                                                AND SUBSTRING(NSGTestDW.dbo.Fact_EventPVR.SerialNumber, 13, 4) = ?
                                                AND NSGTestDW.dbo.Dim_ProcessStep.ProcessStep = ?
                                            GROUP BY
                                                Fact_EventPVR.ComponentEventId
                                                , Fact_EventPVR.SerialNumber
                                                , Dim_DiePVR.CE
                                        )
                                        SELECT
                                            TBL.CE
                                            , TBL.CountCE
                                            , COUNT(DISTINCT TBL.SerialNumber)
                                        FROM
                                            TBL
                                        WHERE
                                            CountCE > 1
                                        GROUP BY
                                            TBL.CE
                                            , TBL.CountCE
                                        ORDER BY
                                            CE""", (row[0], row[1], row[2]))
                    for bow in slursor.fetchall():
                        f.write(f"{bow[0]}(Appears {bow[1]} times in {bow[2]} drives)   ".encode())
                    f.write(b"\n")
                elif tow[0][2] < tow[0][1]:
                    f.write(b"\t\t\t-Fewer CE.\n")
            print(datetime.now())
            f.write(b"\tINVALID FIDS BY PROCESS STEP\n")
            slursor.execute(r"""SELECT
                                    ProcessStep
                                FROM
                                    ##die
                                GROUP BY
                                    ##die.ProdCode
                                    , ##die.SKU
                                    , ##die.ProcessStep
                                HAVING
                                    ##die.ProdCode = ?
                                    AND ##die.SKU = ?
                                ORDER BY
                                    SUM(SerialCount) DESC""", (row[0], row[1]))
            print(datetime.now())
            for bow in slursor.fetchall():
                f.write(f"\t\t{bow[0]}\n".encode())
                mursor.execute(r"""SELECT DISTINCT
                                    Dim_DiePVR.CE
                                    , COUNT(DISTINCT Fact_EventPVR.SerialNumber)
                                FROM Fact_EventPVR
                                    INNER JOIN Dim_ProcessStep
                                        ON Fact_EventPVR.ProcessStepId = Dim_ProcessStep.ProcessStepId
                                    INNER JOIN Dim_DiePVR
                                        ON Dim_DiePVR.ComponentEventId = Fact_EventPVR.ComponentEventId
                                WHERE
                                    Dim_DiePVR.FID = 'data_invalid'
                                GROUP BY
                                    SUBSTRING(Fact_EventPVR.SerialNumber, 3, 2)
                                    , SUBSTRING(Fact_EventPVR.SerialNumber, 13, 4)
                                    , Dim_ProcessStep.ProcessStep
                                    , Dim_DiePVR.CE
                                HAVING
                                    SUBSTRING(Fact_EventPVR.SerialNumber, 3, 2) = ?
                                    AND SUBSTRING(Fact_EventPVR.SerialNumber, 13, 4) = ?
                                    AND Dim_ProcessStep.ProcessStep = ?
                                ORDER BY
                                    CE""", (row[0], row[1], row[2]))
                for slow in mursor.fetchall():
                    f.write(f"\t\t\tCE {slow[0]} has an invalid FID in {slow[1]} drives. ".encode())
                f.write(b"\n")
            print(datetime.now())
            f.write(b"\n")
        print(datetime.now())

if __name__ == '__main__':
    main()
