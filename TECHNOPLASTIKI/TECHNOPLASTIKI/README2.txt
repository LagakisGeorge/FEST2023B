

PARTIDES 
   ' ΒΡΕΣ (pos2)  ΤΟ ΠΡΩΤΟ ΤΙΜΟΛΟΓΙΟ ΠΟΥ ΕΧΕΙ ΥΠΟΛΟΙΠΟ ΜΕ ΑΥΤΟ ΤΟ ΣΥΣΤΑΤΙΚΟ
   SELECT TOP 1 KOD,YPOL,ATIM,HME,ID  FROM TIMS where YPOL>0 AND RTRIM(LTRIM(KOD))='" + YL + "' ORDER BY HME ", pos2)

PALETO : ΙΔΙΟ ΑΚΡΙΒΩΣ ΜΕ PARTIDES ΑΛΛΑ ΑΝΑΛΥΤΙΚΟ
CREATE TABLE [dbo].[PALETO](
	[PARTIDA] [bigint] NULL,
	[HME] [datetime] NULL,
	[KOD] [nvarchar](50) NULL,
	[TIMOLOGIA] [nvarchar](max) NULL,
	[TEMAXIA] [bigint] NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[N1] [int] NULL,
	[N2] [int] NULL,
	[CH1] [nvarchar](50) NULL,
	[CH2] [nvarchar](50) NULL,
	[YPOL] [bigint] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]


'ΑΛΛΑ ΑΝ ΕΧΩ 5 ΠΑΛΕΤΑ/ΕΤΙΚΕΤΤΕΣ ΣΤΗΝ ΙΔΙΑ ΠΑΡΤΙΔΑ ΘΑ ΕΧΩ 5 ΕΓΓΡΑΦΕΣ ΣΤΟ PALETO KAI 1 ΕΓΓΡΑΦΗ ΣΤΟ PARTIDES
            



TIMS  : ΤΙΜΟΛΟΓΙΑ ΑΓΟΡΑΣ


TIMSANAL :ΚΙΝΗΣΕΙΣ  ΤΙΜΟΛΟΓΙΩΝ



TIMSPOL  :TIMOLOGIA ΠΩΛΗΣΗΣ

