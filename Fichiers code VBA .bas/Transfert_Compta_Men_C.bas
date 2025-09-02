Attribute VB_Name = "Transfert_Compta_Men_C"



Sub Transferts_Mensuels_C()


Windows("Mensuel.xlsx").Activate
Sheets("C").Select
    Call Transfert_Mensuel_Janvier_C
    Call Transfert_Mensuel_Février_C
    Call Transfert_Mensuel_Mars_C
    Call Transfert_Mensuel_Avril_C
    Call Transfert_Mensuel_Mai_C
    Call Transfert_Mensuel_Juin_C
    Call Transfert_Mensuel_Juillet_C
    'Call Transfert_Mensuel_Août_C
    'Call Transfert_Mensuel_Septembre_C
    'Call Transfert_Mensuel_Octobre_C
    'Call Transfert_Mensuel_Novembre_C
    'Call Transfert_Mensuel_Décembre_C


End Sub


Sub Transfert_Mensuel_Janvier_C()


    Range("GH12").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R7C190"
    Range("GP12").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R68C198"
    Range("GT12").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R68C202"
    
    Range("GH13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R75C190"
    Range("GP13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R136C198"
    Range("GT13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R136C202"
    
    Range("GH14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R143C190"
    Range("GP14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R204C198"
    Range("GT14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R204C202"
    
    Range("GH15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R211C190"
    Range("GP15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R272C198"
    Range("GT15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R272C202"
    
    Range("GH16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R279C190"
    Range("GP16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R340C198"
    Range("GT16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R340C202"
    
    Range("GH17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R347C190"
    Range("GP17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R408C198"
    Range("GT17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R408C202"
    
    
    Range("GH18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R415C190"
    Range("GP18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R476C198"
    Range("GT18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R476C202"
    
    Range("GH19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R483C190"
    Range("GP19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R544C198"
    Range("GT19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R544C202"
    
    Range("GH20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R551C190"
    Range("GP20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R612C198"
    Range("GT20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R612C202"
    
    Range("GH21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R619C190"
    Range("GP21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R680C198"
    Range("GT21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R680C202"
    
    Range("GH22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R687C190"
    Range("GP22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R748C198"
    Range("GT22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R748C202"
    
    Range("GH23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C190"
    Range("GP23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R816C198"
    Range("GT23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R816C202"
    
    Range("GH24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R823C190"
    Range("GP24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R884C198"
    Range("GT24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R884C202"
    
    Range("GH25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R891C190"
    Range("GP25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R952C198"
    Range("GT25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R952C202"
    
    Range("GH26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R959C190"
    Range("GP26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1020C198"
    Range("GT26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1020C202"
    
    Range("GH27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1027C190"
    Range("GP27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1088C198"
    Range("GT27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1088C202"
    
    Range("GH28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1095C190"
    Range("GP28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1156C198"
    Range("GT28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1156C202"
    
    Range("GH29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1163C190"
    Range("GP29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1224C198"
    Range("GT29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1224C202"
    
    Range("GH30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1231C190"
    Range("GP30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1292C198"
    Range("GT30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1292C202"
    
    Range("GH31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1299C190"
    Range("GP31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1360C198"
    Range("GT31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1360C202"
    
    Range("GH32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1367C190"
    Range("GP32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1428C198"
    Range("GT32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1428C202"
    
    Range("GH33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1435C190"
    Range("GP33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1496C198"
    Range("GT33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1496C202"
    
    Range("GH34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1503C190"
    Range("GP34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1564C198"
    Range("GT34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1564C202"
    
    Range("GH35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1571C190"
    Range("GP35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1632C198"
    Range("GT35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1632C202"
    
    Range("GH36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1639C190"
    Range("GP36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C198"
    Range("GT36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C202"
    
    Range("GH37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1707C190"
    Range("GP37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1768C198"
    Range("GT37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1768C202"
    
    Range("GH38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1775C190"
    Range("GP38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1836C198"
    Range("GT38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1836C202"
    
    Range("GH39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1843C190"
    Range("GP39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1904C198"
    Range("GT39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1904C202"
    
    Range("GH40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1911C190"
    Range("GP40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1972C198"
    Range("GT40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1972C202"
    
    Range("GH41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1979C190"
    Range("GP41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2040C198"
    Range("GT41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2040C202"
    
    Range("GH42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2047C190"
    Range("GP42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2108C198"
    Range("GT42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2108C202"
    
    
    
    
    Range("GH43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2115C190"
    Range("GP43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2176C198"
    Range("GT43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2176C202"
    
    Range("GH44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2183C190"
    Range("GP44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2244C198"
    Range("GT44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2244C202"
    
    Range("GH45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2251C190"
    Range("GP45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2312C198"
    Range("GT45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2312C202"
    
    Range("GH46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2319C190"
    Range("GP46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2380C198"
    Range("GT46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2380C202"
    
    
    
    
    Range("GH47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2387C190"
    Range("GP47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2448C198"
    Range("GT47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2448C202"
    
    Range("GH48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2455C190"
    Range("GP48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2516C198"
    Range("GT48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2516C202"
    
    Range("GH49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2523C190"
    Range("GP49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2584C198"
    Range("GT49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2584C202"
    
    Range("GH50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C190"
    Range("GP50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2652C198"
    Range("GT50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2652C202"
    
    
    
    
    
    Range("GH51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2659C190"
    Range("GP51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2720C198"
    Range("GT51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2720C202"
    
    Range("GH52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2727C190"
    Range("GP52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2788C198"
    Range("GT52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2788C202"
    
    Range("GH53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2795C190"
    Range("GP53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2856C198"
    Range("GT53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2856C202"
    
    Range("GH54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2863C190"
    Range("GP54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2924C198"
    Range("GT54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2924C202"
    
    
    
    
    
    Range("GH55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2931C190"
    Range("GP55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2992C198"
    Range("GT55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2992C202"
    
    Range("GH56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2999C190"
    Range("GP56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R3060C198"
    Range("GT56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R3060C202"
    
    Range("GH57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R3067C190"
    Range("GP57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R3128C198"
    Range("GT57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R3128C202"
    
   
    
        
End Sub


Sub Transfert_Mensuel_Février_C()


    Range("FQ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C173"
    Range("FY13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C181"
    Range("GC13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C185"
    Range("FQ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C173"
    Range("FY14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C181"
    Range("GC14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C185"
    Range("FQ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C173"
    Range("FY15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C181"
    Range("GC15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C185"
    Range("FQ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C173"
    Range("FY16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C181"
    Range("GC16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C185"
    Range("FQ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C173"
    Range("FY17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C181"
    Range("GC17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C185"
    Range("FQ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C173"
    Range("FY18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C181"
    Range("GC18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C185"
    Range("FQ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C173"
    Range("FY19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C181"
    Range("GC19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C185"
    Range("FQ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C173"
    Range("FY20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C181"
    Range("GC20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C185"
    Range("FQ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C173"
    Range("FY21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C181"
    Range("GC21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C185"
    Range("FQ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C173"
    Range("FY22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C181"
    Range("GC22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C185"
    Range("FQ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C173"
    Range("FY23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C181"
    Range("GC23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C185"
    Range("FQ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C173"
    Range("FY24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C181"
    Range("GC24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C185"
    Range("FQ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C173"
    Range("FY25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C181"
    Range("GC25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C185"
    Range("FQ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C173"
    Range("FY26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C181"
    Range("GC26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C185"
    Range("FQ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C173"
    Range("FY27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C181"
    Range("GC27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C185"
    Range("FQ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C173"
    Range("FY28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C181"
    Range("GC28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C185"
    Range("FQ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C173"
    Range("FY29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C181"
    Range("GC29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C185"
    Range("FQ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C173"
    Range("FY30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C181"
    Range("GC30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C185"
    Range("FQ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C173"
    Range("FY31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C181"
    Range("GC31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C185"
    Range("FQ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C173"
    Range("FY32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C181"
    Range("GC32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C185"
    Range("FQ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C173"
    Range("FY33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C181"
    Range("GC33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C185"
    Range("FQ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C173"
    Range("FY34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C181"
    Range("GC34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C185"
    Range("FQ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C173"
    Range("FY35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C181"
    Range("GC35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C185"
    Range("FQ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C173"
    Range("FY36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C181"
    Range("GC36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C185"
    Range("FQ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C173"
    Range("FY37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C181"
    Range("GC37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C185"
    Range("FQ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C173"
    Range("FY38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C181"
    Range("GC38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C185"
    Range("FQ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C173"
    Range("FY39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C181"
    Range("GC39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C185"
    Range("FQ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C173"
    Range("FY40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C181"
    Range("GC40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C185"
    Range("FQ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C173"
    Range("FY41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C181"
    Range("GC41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C185"
    Range("FQ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C173"
    Range("FY42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C181"
    Range("GC42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C185"
    Range("FQ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C173"
    Range("FY43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C181"
    Range("GC43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C185"
    Range("FQ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C173"
    Range("FY44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C181"
    Range("GC44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C185"
    Range("FQ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C173"
    Range("FY45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C181"
    Range("GC45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C185"
    Range("FQ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C173"
    Range("FY46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C181"
    Range("GC46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C185"
    Range("FQ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C173"
    Range("FY47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C181"
    Range("GC47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C185"
    Range("FQ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C173"
    Range("FY48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C181"
    Range("GC48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C185"
    Range("FQ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C173"
    Range("FY49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C181"
    Range("GC49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C185"
    Range("FQ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C173"
    Range("FY50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C181"
    Range("GC50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C185"
    Range("FQ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C173"
    Range("FY51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C181"
    Range("GC51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C185"
    Range("FQ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C173"
    Range("FY52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C181"
    Range("GC52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C185"
    Range("FQ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C173"
    Range("FY53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C181"
    Range("GC53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C185"
    Range("FQ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C173"
    Range("FY54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C181"
    Range("GC54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C185"
    Range("FQ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C173"
    Range("FY55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C181"
    Range("GC55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C185"
    Range("FQ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C173"
    Range("FY56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C181"
    Range("GC56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C185"
    Range("FQ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C173"
    Range("FY57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C181"
    Range("GC57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C185"
    Range("FQ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C173"
    Range("FY58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C181"
    Range("GC58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C185"
    
    

End Sub


Sub Transfert_Mensuel_Mars_C()


    Range("EZ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C156"
    Range("FH13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C164"
    Range("FL13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C168"
    Range("EZ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C156"
    Range("FH14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C164"
    Range("FL14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C168"
    Range("EZ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C156"
    Range("FH15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C164"
    Range("FL15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C168"
    Range("EZ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C156"
    Range("FH16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C164"
    Range("FL16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C168"
    Range("EZ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C156"
    Range("FH17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C164"
    Range("FL17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C168"
    Range("EZ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C156"
    Range("FH18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C164"
    Range("FL18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C168"
    Range("EZ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C156"
    Range("FH19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C164"
    Range("FL19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C168"
    Range("EZ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C156"
    Range("FH20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C164"
    Range("FL20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C168"
    Range("EZ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C156"
    Range("FH21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C164"
    Range("FL21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C168"
    Range("EZ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C156"
    Range("FH22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C164"
    Range("FL22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C168"
    Range("EZ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C156"
    Range("FH23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C164"
    Range("FL23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C168"
    Range("EZ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C156"
    Range("FH24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C164"
    Range("FL24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C168"
    Range("EZ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C156"
    Range("FH25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C164"
    Range("FL25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C168"
    Range("EZ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C156"
    Range("FH26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C164"
    Range("FL26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C168"
    Range("EZ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C156"
    Range("FH27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C164"
    Range("FL27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C168"
    Range("EZ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C156"
    Range("FH28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C164"
    Range("FL28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C168"
    Range("EZ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C156"
    Range("FH29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C164"
    Range("FL29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C168"
    Range("EZ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C156"
    Range("FH30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C164"
    Range("FL30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C168"
    Range("EZ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C156"
    Range("FH31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C164"
    Range("FL31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C168"
    Range("EZ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C156"
    Range("FH32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C164"
    Range("FL32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C168"
    Range("EZ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C156"
    Range("FH33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C164"
    Range("FL33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C168"
    Range("EZ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C156"
    Range("FH34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C164"
    Range("FL34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C168"
    Range("EZ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C156"
    Range("FH35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C164"
    Range("FL35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C168"
    Range("EZ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C156"
    Range("FH36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C164"
    Range("FL36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C168"
    Range("EZ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C156"
    Range("FH37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C164"
    Range("FL37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C168"
    Range("EZ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C156"
    Range("FH38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C164"
    Range("FL38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C168"
    Range("EZ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C156"
    Range("FH39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C164"
    Range("FL39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C168"
    Range("EZ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C156"
    Range("FH40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C164"
    Range("FL40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C168"
    Range("EZ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C156"
    Range("FH41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C164"
    Range("FL41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C168"
    Range("EZ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C156"
    Range("FH42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C164"
    Range("FL42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C168"
    Range("EZ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C156"
    Range("FH43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C164"
    Range("FL43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C168"
    Range("EZ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C156"
    Range("FH44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C164"
    Range("FL44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C168"
    Range("EZ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C156"
    Range("FH45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C164"
    Range("FL45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C168"
    Range("EZ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C156"
    Range("FH46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C164"
    Range("FL46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C168"
    Range("EZ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C156"
    Range("FH47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C164"
    Range("FL47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C168"
    Range("EZ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C156"
    Range("FH48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C164"
    Range("FL48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C168"
    Range("EZ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C156"
    Range("FH49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C164"
    Range("FL49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C168"
    Range("EZ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C156"
    Range("FH50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C164"
    Range("FL50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C168"
    Range("EZ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C156"
    Range("FH51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C164"
    Range("FL51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C168"
    Range("EZ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C156"
    Range("FH52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C164"
    Range("FL52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C168"
    Range("EZ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C156"
    Range("FH53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C164"
    Range("FL53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C168"
    Range("EZ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C156"
    Range("FH54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C164"
    Range("FL54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C168"
    Range("EZ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C156"
    Range("FH55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C164"
    Range("FL55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C168"
    Range("EZ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C156"
    Range("FH56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C164"
    Range("FL56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C168"
    Range("EZ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C156"
    Range("FH57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C164"
    Range("FL57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C168"
    Range("EZ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C156"
    Range("FH58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C164"
    Range("FL58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C168"
    
    
End Sub


Sub Transfert_Mensuel_Avril_C()


    Range("EI13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C139"
    Range("EQ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C147"
    Range("EU13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C151"
    Range("EI14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C139"
    Range("EQ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C147"
    Range("EU14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C151"
    Range("EI15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C139"
    Range("EQ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C147"
    Range("EU15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C151"
    Range("EI16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C139"
    Range("EQ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C147"
    Range("EU16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C151"
    Range("EI17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C139"
    Range("EQ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C147"
    Range("EU17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C151"
    Range("EI18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C139"
    Range("EQ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C147"
    Range("EU18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C151"
    Range("EI19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C139"
    Range("EQ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C147"
    Range("EU19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C151"
    Range("EI20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C139"
    Range("EQ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C147"
    Range("EU20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C151"
    Range("EI21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C139"
    Range("EQ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C147"
    Range("EU21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C151"
    Range("EI22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C139"
    Range("EQ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C147"
    Range("EU22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C151"
    Range("EI23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C139"
    Range("EQ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C147"
    Range("EU23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C151"
    Range("EI24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C139"
    Range("EQ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C147"
    Range("EU24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C151"
    Range("EI25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C139"
    Range("EQ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C147"
    Range("EU25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C151"
    Range("EI26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C139"
    Range("EQ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C147"
    Range("EU26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C151"
    Range("EI27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C139"
    Range("EQ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C147"
    Range("EU27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C151"
    Range("EI28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C139"
    Range("EQ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C147"
    Range("EU28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C151"
    Range("EI29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C139"
    Range("EQ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C147"
    Range("EU29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C151"
    Range("EI30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C139"
    Range("EQ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C147"
    Range("EU30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C151"
    Range("EI31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C139"
    Range("EQ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C147"
    Range("EU31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C151"
    Range("EI32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C139"
    Range("EQ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C147"
    Range("EU32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C151"
    Range("EI33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C139"
    Range("EQ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C147"
    Range("EU33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C151"
    Range("EI34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C139"
    Range("EQ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C147"
    Range("EU34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C151"
    Range("EI35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C139"
    Range("EQ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C147"
    Range("EU35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C151"
    Range("EI36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C139"
    Range("EQ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C147"
    Range("EU36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C151"
    Range("EI37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C139"
    Range("EQ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C147"
    Range("EU37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C151"
    Range("EI38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C139"
    Range("EQ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C147"
    Range("EU38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C151"
    Range("EI39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C139"
    Range("EQ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C147"
    Range("EU39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C151"
    Range("EI40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C139"
    Range("EQ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C147"
    Range("EU40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C151"
    Range("EI41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C139"
    Range("EQ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C147"
    Range("EU41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C151"
    Range("EI42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C139"
    Range("EQ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C147"
    Range("EU42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C151"
    Range("EI43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C139"
    Range("EQ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C147"
    Range("EU43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C151"
    Range("EI44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C139"
    Range("EQ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C147"
    Range("EU44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C151"
    Range("EI45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C139"
    Range("EQ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C147"
    Range("EU45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C151"
    Range("EI46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C139"
    Range("EQ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C147"
    Range("EU46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C151"
    Range("EI47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C139"
    Range("EQ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C147"
    Range("EU47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C151"
    Range("EI48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C139"
    Range("EQ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C147"
    Range("EU48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C151"
    Range("EI49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C139"
    Range("EQ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C147"
    Range("EU49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C151"
    Range("EI50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C139"
    Range("EQ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C147"
    Range("EU50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C151"
    Range("EI51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C139"
    Range("EQ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C147"
    Range("EU51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C151"
    Range("EI52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C139"
    Range("EQ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C147"
    Range("EU52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C151"
    Range("EI53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C139"
    Range("EQ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C147"
    Range("EU53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C151"
    Range("EI54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C139"
    Range("EQ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C147"
    Range("EU54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C151"
    Range("EI55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C139"
    Range("EQ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C147"
    Range("EU55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C151"
    Range("EI56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C139"
    Range("EQ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C147"
    Range("EU56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C151"
    Range("EI57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C139"
    Range("EQ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C147"
    Range("EU57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C151"
    Range("EI58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C139"
    Range("EQ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C147"
    Range("EU58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C151"


End Sub


Sub Transfert_Mensuel_Mai_C()


    Range("DR13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C122"
    Range("DZ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C130"
    Range("ED13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C134"
    Range("DR14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C122"
    Range("DZ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C130"
    Range("ED14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C134"
    Range("DR15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C122"
    Range("DZ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C130"
    Range("ED15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C134"
    Range("DR16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C122"
    Range("DZ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C130"
    Range("ED16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C134"
    Range("DR17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C122"
    Range("DZ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C130"
    Range("ED17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C134"
    Range("DR18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C122"
    Range("DZ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C130"
    Range("ED18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C134"
    Range("DR19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C122"
    Range("DZ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C130"
    Range("ED19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C134"
    Range("DR20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C122"
    Range("DZ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C130"
    Range("ED20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C134"
    Range("DR21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C122"
    Range("DZ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C130"
    Range("ED21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C134"
    Range("DR22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C122"
    Range("DZ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C130"
    Range("ED22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C134"
    Range("DR23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C122"
    Range("DZ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C130"
    Range("ED23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C134"
    Range("DR24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C122"
    Range("DZ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C130"
    Range("ED24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C134"
    Range("DR25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C122"
    Range("DZ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C130"
    Range("ED25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C134"
    Range("DR26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C122"
    Range("DZ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C130"
    Range("ED26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C134"
    Range("DR27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C122"
    Range("DZ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C130"
    Range("ED27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C134"
    Range("DR28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C122"
    Range("DZ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C130"
    Range("ED28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C134"
    Range("DR29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C122"
    Range("DZ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C130"
    Range("ED29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C134"
    Range("DR30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C122"
    Range("DZ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C130"
    Range("ED30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C134"
    Range("DR31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C122"
    Range("DZ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C130"
    Range("ED31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C134"
    Range("DR32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C122"
    Range("DZ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C130"
    Range("ED32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C134"
    Range("DR33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C122"
    Range("DZ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C130"
    Range("ED33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C134"
    Range("DR34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C122"
    Range("DZ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C130"
    Range("ED34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C134"
    Range("DR35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C122"
    Range("DZ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C130"
    Range("ED35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C134"
    Range("DR36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C122"
    Range("DZ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C130"
    Range("ED36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C134"
    Range("DR37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C122"
    Range("DZ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C130"
    Range("ED37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C134"
    Range("DR38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C122"
    Range("DZ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C130"
    Range("ED38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C134"
    Range("DR39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C122"
    Range("DZ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C130"
    Range("ED39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C134"
    Range("DR40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C122"
    Range("DZ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C130"
    Range("ED40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C134"
    Range("DR41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C122"
    Range("DZ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C130"
    Range("ED41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C134"
    Range("DR42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C122"
    Range("DZ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C130"
    Range("ED42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C134"
    Range("DR43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C122"
    Range("DZ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C130"
    Range("ED43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C134"
    Range("DR44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C122"
    Range("DZ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C130"
    Range("ED44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C134"
    Range("DR45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C122"
    Range("DZ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C130"
    Range("ED45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C134"
    Range("DR46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C122"
    Range("DZ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C130"
    Range("ED46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C134"
    Range("DR47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C122"
    Range("DZ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C130"
    Range("ED47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C134"
    Range("DR48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C122"
    Range("DZ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C130"
    Range("ED48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C134"
    Range("DR49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C122"
    Range("DZ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C130"
    Range("ED49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C134"
    Range("DR50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C122"
    Range("DZ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C130"
    Range("ED50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C134"
    Range("DR51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C122"
    Range("DZ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C130"
    Range("ED51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C134"
    Range("DR52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C122"
    Range("DZ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C130"
    Range("ED52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C134"
    Range("DR53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C122"
    Range("DZ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C130"
    Range("ED53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C134"
    Range("DR54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C122"
    Range("DZ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C130"
    Range("ED54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C134"
    Range("DR55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C122"
    Range("DZ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C130"
    Range("ED55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C134"
    Range("DR56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C122"
    Range("DZ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C130"
    Range("ED56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C134"
    Range("DR57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C122"
    Range("DZ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C130"
    Range("ED57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C134"
    Range("DR58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C122"
    Range("DZ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C130"
    Range("ED58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C134"
        
   
End Sub


Sub Transfert_Mensuel_Juin_C()


    Range("DA13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C105"
    Range("DI13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C113"
    Range("DM13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C117"
    Range("DA14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C105"
    Range("DI14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C113"
    Range("DM14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C117"
    Range("DA15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C105"
    Range("DI15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C113"
    Range("DM15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C117"
    Range("DA16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C105"
    Range("DI16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C113"
    Range("DM16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C117"
    Range("DA17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C105"
    Range("DI17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C113"
    Range("DM17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C117"
    Range("DA18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C105"
    Range("DI18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C113"
    Range("DM18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C117"
    Range("DA19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C105"
    Range("DI19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C113"
    Range("DM19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C117"
    Range("DA20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C105"
    Range("DI20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C113"
    Range("DM20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C117"
    Range("DA21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C105"
    Range("DI21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C113"
    Range("DM21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C117"
    Range("DA22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C105"
    Range("DI22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C113"
    Range("DM22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C117"
    Range("DA23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C105"
    Range("DI23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C113"
    Range("DM23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C117"
    Range("DA24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C105"
    Range("DI24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C113"
    Range("DM24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C117"
    Range("DA25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C105"
    Range("DI25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C113"
    Range("DM25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C117"
    Range("DA26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C105"
    Range("DI26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C113"
    Range("DM26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C117"
    Range("DA27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C105"
    Range("DI27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C113"
    Range("DM27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C117"
    Range("DA28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C105"
    Range("DI28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C113"
    Range("DM28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C117"
    Range("DA29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C105"
    Range("DI29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C113"
    Range("DM29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C117"
    Range("DA30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C105"
    Range("DI30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C113"
    Range("DM30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C117"
    Range("DA31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C105"
    Range("DI31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C113"
    Range("DM31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C117"
    Range("DA32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C105"
    Range("DI32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C113"
    Range("DM32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C117"
    Range("DA33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C105"
    Range("DI33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C113"
    Range("DM33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C117"
    Range("DA34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C105"
    Range("DI34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C113"
    Range("DM34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C117"
    Range("DA35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C105"
    Range("DI35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C113"
    Range("DM35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C117"
    Range("DA36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C105"
    Range("DI36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C113"
    Range("DM36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C117"
    Range("DA37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C105"
    Range("DI37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C113"
    Range("DM37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C117"
    Range("DA38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C105"
    Range("DI38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C113"
    Range("DM38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C117"
    Range("DA39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C105"
    Range("DI39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C113"
    Range("DM39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C117"
    Range("DA40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C105"
    Range("DI40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C113"
    Range("DM40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C117"
    Range("DA41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C105"
    Range("DI41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C113"
    Range("DM41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C117"
    Range("DA42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C105"
    Range("DI42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C113"
    Range("DM42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C117"
    Range("DA43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C105"
    Range("DI43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C113"
    Range("DM43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C117"
    Range("DA44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C105"
    Range("DI44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C113"
    Range("DM44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C117"
    Range("DA45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C105"
    Range("DI45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C113"
    Range("DM45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C117"
    Range("DA46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C105"
    Range("DI46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C113"
    Range("DM46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C117"
    Range("DA47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C105"
    Range("DI47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C113"
    Range("DM47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C117"
    Range("DA48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C105"
    Range("DI48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C113"
    Range("DM48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C117"
    Range("DA49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C105"
    Range("DI49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C113"
    Range("DM49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C117"
    Range("DA50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C105"
    Range("DI50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C113"
    Range("DM50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C117"
    Range("DA51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C105"
    Range("DI51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C113"
    Range("DM51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C117"
    Range("DA52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C105"
    Range("DI52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C113"
    Range("DM52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C117"
    Range("DA53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C105"
    Range("DI53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C113"
    Range("DM53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C117"
    Range("DA54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C105"
    Range("DI54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C113"
    Range("DM54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C117"
    Range("DA55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C105"
    Range("DI55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C113"
    Range("DM55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C117"
    Range("DA56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C105"
    Range("DI56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C113"
    Range("DM56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C117"
    Range("DA57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C105"
    Range("DI57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C113"
    Range("DM57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C117"
    Range("DA58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C105"
    Range("DI58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C113"
    Range("DM58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C117"
    
 
End Sub


Sub Transfert_Mensuel_Juillet_C()


    Range("CJ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R8C88"
    Range("CR13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C96"
    Range("CV13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R62C100"
    Range("CJ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R71C88"
    Range("CR14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C96"
    Range("CV14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R125C100"
    Range("CJ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R134C88"
    Range("CR15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C96"
    Range("CV15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R188C100"
    Range("CJ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R197C88"
    Range("CR16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C96"
    Range("CV16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R251C100"
    Range("CJ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R260C88"
    Range("CR17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C96"
    Range("CV17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R314C100"
    Range("CJ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R323C88"
    Range("CR18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C96"
    Range("CV18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R377C100"
    Range("CJ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R386C88"
    Range("CR19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C96"
    Range("CV19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R440C100"
    Range("CJ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R449C88"
    Range("CR20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C96"
    Range("CV20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R503C100"
    Range("CJ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R512C88"
    Range("CR21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C96"
    Range("CV21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R566C100"
    Range("CJ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R575C88"
    Range("CR22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C96"
    Range("CV22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R629C100"
    Range("CJ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R638C88"
    Range("CR23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C96"
    Range("CV23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R692C100"
    Range("CJ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R701C88"
    Range("CR24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C96"
    Range("CV24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R755C100"
    Range("CJ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R764C88"
    Range("CR25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C96"
    Range("CV25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R818C100"
    Range("CJ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R827C88"
    Range("CR26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C96"
    Range("CV26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R881C100"
    Range("CJ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R890C88"
    Range("CR27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C96"
    Range("CV27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R944C100"
    Range("CJ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R953C88"
    Range("CR28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C96"
    Range("CV28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1007C100"
    Range("CJ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1016C88"
    Range("CR29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C96"
    Range("CV29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1070C100"
    Range("CJ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1079C88"
    Range("CR30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C96"
    Range("CV30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1133C100"
    Range("CJ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1142C88"
    Range("CR31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C96"
    Range("CV31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1196C100"
    Range("CJ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1205C88"
    Range("CR32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C96"
    Range("CV32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1259C100"
    Range("CJ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1268C88"
    Range("CR33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C96"
    Range("CV33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1322C100"
    Range("CJ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1331C88"
    Range("CR34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C96"
    Range("CV34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1385C100"
    Range("CJ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1394C88"
    Range("CR35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C96"
    Range("CV35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1448C100"
    Range("CJ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1457C88"
    Range("CR36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C96"
    Range("CV36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1511C100"
    Range("CJ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1520C88"
    Range("CR37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C96"
    Range("CV37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1574C100"
    Range("CJ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1583C88"
    Range("CR38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C96"
    Range("CV38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1637C100"
    Range("CJ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1646C88"
    Range("CR39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C96"
    Range("CV39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1700C100"
    Range("CJ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1709C88"
    Range("CR40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C96"
    Range("CV40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1763C100"
    Range("CJ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1772C88"
    Range("CR41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C96"
    Range("CV41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1826C100"
    Range("CJ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1835C88"
    Range("CR42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C96"
    Range("CV42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1889C100"
    Range("CJ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1898C88"
    Range("CR43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C96"
    Range("CV43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1952C100"
    Range("CJ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R1961C88"
    Range("CR44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C96"
    Range("CV44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2015C100"
    Range("CJ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2024C88"
    Range("CR45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C96"
    Range("CV45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2078C100"
    Range("CJ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2087C88"
    Range("CR46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C96"
    Range("CV46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2141C100"
    Range("CJ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2150C88"
    Range("CR47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C96"
    Range("CV47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2204C100"
    Range("CJ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2213C88"
    Range("CR48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C96"
    Range("CV48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2267C100"
    Range("CJ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2276C88"
    Range("CR49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C96"
    Range("CV49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2330C100"
    Range("CJ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2339C88"
    Range("CR50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C96"
    Range("CV50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2393C100"
    Range("CJ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2402C88"
    Range("CR51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C96"
    Range("CV51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2456C100"
    Range("CJ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2465C88"
    Range("CR52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C96"
    Range("CV52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2519C100"
    Range("CJ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2528C88"
    Range("CR53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C96"
    Range("CV53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2582C100"
    Range("CJ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2591C88"
    Range("CR54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C96"
    Range("CV54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2645C100"
    Range("CJ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2654C88"
    Range("CR55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C96"
    Range("CV55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2708C100"
    Range("CJ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2717C88"
    Range("CR56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C96"
    Range("CV56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2771C100"
    Range("CJ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2780C88"
    Range("CR57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C96"
    Range("CV57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2834C100"
    Range("CJ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2843C88"
    Range("CR58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C96"
    Range("CV58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xlsx]C!R2897C100"
    
   
End Sub


Sub Transfert_Mensuel_Août_C()


    Range("BS13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R8C71"
    Range("CA13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C79"
    Range("CE13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C83"
    Range("BS14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R71C71"
    Range("CA14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C79"
    Range("CE14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C83"
    Range("BS15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R134C71"
    Range("CA15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C79"
    Range("CE15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C83"
    Range("BS16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R197C71"
    Range("CA16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C79"
    Range("CE16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C83"
    Range("BS17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R260C71"
    Range("CA17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C79"
    Range("CE17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C83"
    Range("BS18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R323C71"
    Range("CA18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C79"
    Range("CE18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C83"
    Range("BS19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R386C71"
    Range("CA19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C79"
    Range("CE19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C83"
    Range("BS20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R449C71"
    Range("CA20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C79"
    Range("CE20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C83"
    Range("BS21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R512C71"
    Range("CA21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C79"
    Range("CE21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C83"
    Range("BS22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R575C71"
    Range("CA22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C79"
    Range("CE22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C83"
    Range("BS23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R638C71"
    Range("CA23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C79"
    Range("CE23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C83"
    Range("BS24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R701C71"
    Range("CA24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C79"
    Range("CE24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C83"
    Range("BS25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R764C71"
    Range("CA25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C79"
    Range("CE25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C83"
    Range("BS26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R827C71"
    Range("CA26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C79"
    Range("CE26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C83"
    Range("BS27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R890C71"
    Range("CA27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C79"
    Range("CE27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C83"
    Range("BS28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R953C71"
    Range("CA28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C79"
    Range("CE28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C83"
    Range("BS29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1016C71"
    Range("CA29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C79"
    Range("CE29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C83"
    Range("BS30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1079C71"
    Range("CA30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C79"
    Range("CE30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C83"
    Range("BS31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1142C71"
    Range("CA31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C79"
    Range("CE31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C83"
    Range("BS32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1205C71"
    Range("CA32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C79"
    Range("CE32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C83"
    Range("BS33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1268C71"
    Range("CA33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C79"
    Range("CE33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C83"
    Range("BS34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1331C71"
    Range("CA34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C79"
    Range("CE34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C83"
    Range("BS35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1394C71"
    Range("CA35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C79"
    Range("CE35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C83"
    Range("BS36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1457C71"
    Range("CA36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C79"
    Range("CE36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C83"
    Range("BS37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1520C71"
    Range("CA37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C79"
    Range("CE37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C83"
    Range("BS38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1583C71"
    Range("CA38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C79"
    Range("CE38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C83"
    Range("BS39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1646C71"
    Range("CA39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C79"
    Range("CE39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C83"
    Range("BS40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1709C71"
    Range("CA40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C79"
    Range("CE40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C83"
    Range("BS41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1772C71"
    Range("CA41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C79"
    Range("CE41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C83"
    Range("BS42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1835C71"
    Range("CA42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C79"
    Range("CE42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C83"
    Range("BS43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1898C71"
    Range("CA43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C79"
    Range("CE43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C83"
    Range("BS44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1961C71"
    Range("CA44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C79"
    Range("CE44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C83"
    Range("BS45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2024C71"
    Range("CA45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C79"
    Range("CE45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C83"
    Range("BS46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2087C71"
    Range("CA46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C79"
    Range("CE46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C83"
    Range("BS47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2150C71"
    Range("CA47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C79"
    Range("CE47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C83"
    Range("BS48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2213C71"
    Range("CA48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C79"
    Range("CE48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C83"
    Range("BS49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2276C71"
    Range("CA49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C79"
    Range("CE49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C83"
    Range("BS50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2339C71"
    Range("CA50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C79"
    Range("CE50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C83"
    Range("BS51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2402C71"
    Range("CA51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C79"
    Range("CE51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C83"
    Range("BS52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2465C71"
    Range("CA52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C79"
    Range("CE52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C83"
    Range("BS53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2528C71"
    Range("CA53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C79"
    Range("CE53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C83"
    Range("BS54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2591C71"
    Range("CA54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C79"
    Range("CE54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C83"
    Range("BS55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2654C71"
    Range("CA55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C79"
    Range("CE55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C83"
    Range("BS56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2717C71"
    Range("CA56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C79"
    Range("CE56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C83"
    Range("BS57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2780C71"
    Range("CA57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C79"
    Range("CE57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C83"
    Range("BS58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2843C71"
    Range("CA58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C79"
    Range("CE58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C83"
    
    
End Sub


Sub Transfert_Mensuel_Septembre_C()


    Range("BB13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R8C54"
    Range("BJ13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C62"
    Range("BN13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C66"
    Range("BB14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R71C54"
    Range("BJ14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C62"
    Range("BN14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C66"
    Range("BB15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R134C54"
    Range("BJ15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C62"
    Range("BN15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C66"
    Range("BB16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R197C54"
    Range("BJ16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C62"
    Range("BN16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C66"
    Range("BB17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R260C54"
    Range("BJ17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C62"
    Range("BN17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C66"
    Range("BB18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R323C54"
    Range("BJ18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C62"
    Range("BN18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C66"
    Range("BB19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R386C54"
    Range("BJ19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C62"
    Range("BN19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C66"
    Range("BB20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R449C54"
    Range("BJ20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C62"
    Range("BN20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C66"
    Range("BB21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R512C54"
    Range("BJ21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C62"
    Range("BN21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C66"
    Range("BB22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R575C54"
    Range("BJ22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C62"
    Range("BN22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C66"
    Range("BB23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R638C54"
    Range("BJ23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C62"
    Range("BN23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C66"
    Range("BB24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R701C54"
    Range("BJ24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C62"
    Range("BN24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C66"
    Range("BB25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R764C54"
    Range("BJ25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C62"
    Range("BN25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C66"
    Range("BB26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R827C54"
    Range("BJ26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C62"
    Range("BN26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C66"
    Range("BB27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R890C54"
    Range("BJ27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C62"
    Range("BN27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C66"
    Range("BB28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R953C54"
    Range("BJ28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C62"
    Range("BN28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C66"
    Range("BB29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1016C54"
    Range("BJ29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C62"
    Range("BN29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C66"
    Range("BB30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1079C54"
    Range("BJ30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C62"
    Range("BN30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C66"
    Range("BB31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1142C54"
    Range("BJ31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C62"
    Range("BN31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C66"
    Range("BB32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1205C54"
    Range("BJ32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C62"
    Range("BN32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C66"
    Range("BB33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1268C54"
    Range("BJ33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C62"
    Range("BN33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C66"
    Range("BB34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1331C54"
    Range("BJ34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C62"
    Range("BN34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C66"
    Range("BB35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1394C54"
    Range("BJ35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C62"
    Range("BN35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C66"
    Range("BB36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1457C54"
    Range("BJ36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C62"
    Range("BN36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C66"
    Range("BB37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1520C54"
    Range("BJ37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C62"
    Range("BN37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C66"
    Range("BB38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1583C54"
    Range("BJ38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C62"
    Range("BN38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C66"
    Range("BB39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1646C54"
    Range("BJ39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C62"
    Range("BN39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C66"
    Range("BB40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1709C54"
    Range("BJ40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C62"
    Range("BN40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C66"
    Range("BB41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1772C54"
    Range("BJ41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C62"
    Range("BN41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C66"
    Range("BB42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1835C54"
    Range("BJ42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C62"
    Range("BN42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C66"

    Range("BB43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1898C54"
    Range("BJ43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C62"
    Range("BN43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C66"
    Range("BB44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1961C54"
    Range("BJ44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C62"
    Range("BN44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C66"
    Range("BB45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2024C54"
    Range("BJ45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C62"
    Range("BN45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C66"
    Range("BB46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2087C54"
    Range("BJ46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C62"
    Range("BN46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C66"
    Range("BB47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2150C54"
    Range("BJ47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C62"
    Range("BN47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C66"
    Range("BB48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2213C54"
    Range("BJ48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C62"
    Range("BN48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C66"
    Range("BB49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2276C54"
    Range("BJ49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C62"
    Range("BN49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C66"
    Range("BB50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2339C54"
    Range("BJ50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C62"
    Range("BN50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C66"
    Range("BB51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2402C54"
    Range("BJ51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C62"
    Range("BN51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C66"
    Range("BB52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2465C54"
    Range("BJ52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C62"
    Range("BN52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C66"
    Range("BB53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2528C54"
    Range("BJ53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C62"
    Range("BN53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C66"
    Range("BB54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2591C54"
    Range("BJ54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C62"
    Range("BN54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C66"
    Range("BB55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2654C54"
    Range("BJ55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C62"
    Range("BN55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C66"
    Range("BB56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2717C54"
    Range("BJ56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C62"
    Range("BN56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C66"
    Range("BB57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2780C54"
    Range("BJ57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C62"
    Range("BN57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C66"
    Range("BB58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2843C54"
    Range("BJ58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C62"
    Range("BN58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C66"
    
    
End Sub


Sub Transfert_Mensuel_Octobre_C()


    Range("AK13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R8C37"
    Range("AS13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C45"
    Range("AW13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C49"
    Range("AK14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R71C37"
    Range("AS14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C45"
    Range("AW14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C49"
    Range("AK15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R134C37"
    Range("AS15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C45"
    Range("AW15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C49"
    Range("AK16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R197C37"
    Range("AS16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C45"
    Range("AW16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C49"
    Range("AK17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R260C37"
    Range("AS17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C45"
    Range("AW17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C49"
    Range("AK18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R323C37"
    Range("AS18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C45"
    Range("AW18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C49"
    Range("AK19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R386C37"
    Range("AS19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C45"
    Range("AW19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C49"
    Range("AK20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R449C37"
    Range("AS20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C45"
    Range("AW20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C49"
    Range("AK21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R512C37"
    Range("AS21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C45"
    Range("AW21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C49"
    Range("AK22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R575C37"
    Range("AS22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C45"
    Range("AW22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C49"
    Range("AK23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R638C37"
    Range("AS23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C45"
    Range("AW23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C49"
    Range("AK24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R701C37"
    Range("AS24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C45"
    Range("AW24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C49"
    Range("AK25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R764C37"
    Range("AS25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C45"
    Range("AW25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C49"
    Range("AK26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R827C37"
    Range("AS26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C45"
    Range("AW26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C49"
    Range("AK27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R890C37"
    Range("AS27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C45"
    Range("AW27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C49"
    Range("AK28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R953C37"
    Range("AS28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C45"
    Range("AW28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C49"
    Range("AK29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1016C37"
    Range("AS29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C45"
    Range("AW29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C49"
    Range("AK30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1079C37"
    Range("AS30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C45"
    Range("AW30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C49"
    Range("AK31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1142C37"
    Range("AS31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C45"
    Range("AW31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C49"
    Range("AK32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1205C37"
    Range("AS32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C45"
    Range("AW32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C49"
    Range("AK33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1268C37"
    Range("AS33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C45"
    Range("AW33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C49"
    Range("AK34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1331C37"
    Range("AS34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C45"
    Range("AW34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C49"
    Range("AK35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1394C37"
    Range("AS35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C45"
    Range("AW35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C49"
    Range("AK36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1457C37"
    Range("AS36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C45"
    Range("AW36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C49"
    Range("AK37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1520C37"
    Range("AS37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C45"
    Range("AW37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C49"
    Range("AK38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1583C37"
    Range("AS38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C45"
    Range("AW38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C49"
    Range("AK39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1646C37"
    Range("AS39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C45"
    Range("AW39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C49"
    Range("AK40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1709C37"
    Range("AS40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C45"
    Range("AW40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C49"
    Range("AK41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1772C37"
    Range("AS41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C45"
    Range("AW41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C49"
    Range("AK42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1835C37"
    Range("AS42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C45"
    Range("AW42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C49"
    Range("AK43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1898C37"
    Range("AS43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C45"
    Range("AW43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C49"
    Range("AK44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1961C37"
    Range("AS44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C45"
    Range("AW44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C49"
    Range("AK45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2024C37"
    Range("AS45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C45"
    Range("AW45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C49"
    Range("AK46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2087C37"
    Range("AS46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C45"
    Range("AW46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C49"
    Range("AK47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2150C37"
    Range("AS47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C45"
    Range("AW47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C49"
    Range("AK48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2213C37"
    Range("AS48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C45"
    Range("AW48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C49"
    Range("AK49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2276C37"
    Range("AS49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C45"
    Range("AW49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C49"
    Range("AK50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2339C37"
    Range("AS50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C45"
    Range("AW50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C49"
    Range("AK51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2402C37"
    Range("AS51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C45"
    Range("AW51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C49"
    Range("AK52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2465C37"
    Range("AS52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C45"
    Range("AW52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C49"
    Range("AK53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2528C37"
    Range("AS53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C45"
    Range("AW53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C49"
    Range("AK54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2591C37"
    Range("AS54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C45"
    Range("AW54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C49"
    Range("AK55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2654C37"
    Range("AS55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C45"
    Range("AW55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C49"
    Range("AK56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2717C37"
    Range("AS56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C45"
    Range("AW56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C49"
    Range("AK57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2780C37"
    Range("AS57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C45"
    Range("AW57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C49"
    Range("AK58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2843C37"
    Range("AS58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C45"
    Range("AW58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C49"
    
    
End Sub


Sub Transfert_Mensuel_Novembre_C()


    Range("T13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R8C20"
    Range("AB13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C28"
    Range("AF13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C32"
    Range("T14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R71C20"
    Range("AB14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C28"
    Range("AF14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C32"
    Range("T15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R134C20"
    Range("AB15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C28"
    Range("AF15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C32"
    Range("T16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R197C20"
    Range("AB16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C28"
    Range("AF16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C32"
    Range("T17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R260C20"
    Range("AB17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C28"
    Range("AF17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C32"
    Range("T18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R323C20"
    Range("AB18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C28"
    Range("AF18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C32"
    Range("T19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R386C20"
    Range("AB19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C28"
    Range("AF19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C32"
    Range("T20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R449C20"
    Range("AB20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C28"
    Range("AF20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C32"
    Range("T21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R512C20"
    Range("AB21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C28"
    Range("AF21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C32"
    Range("T22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R575C20"
    Range("AB22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C28"
    Range("AF22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C32"
    Range("T23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R638C20"
    Range("AB23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C28"
    Range("AF23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C32"
    Range("T24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R701C20"
    Range("AB24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C28"
    Range("AF24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C32"
    Range("T25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R764C20"
    Range("AB25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C28"
    Range("AF25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C32"
    Range("T26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R827C20"
    Range("AB26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C28"
    Range("AF26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C32"
    Range("T27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R890C20"
    Range("AB27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C28"
    Range("AF27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C32"
    Range("T28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R953C20"
    Range("AB28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C28"
    Range("AF28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C32"
    Range("T29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1016C20"
    Range("AB29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C28"
    Range("AF29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C32"
    Range("T30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1079C20"
    Range("AB30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C28"
    Range("AF30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C32"
    Range("T31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1142C20"
    Range("AB31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C28"
    Range("AF31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C32"
    Range("T32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1205C20"
    Range("AB32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C28"
    Range("AF32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C32"
    Range("T33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1268C20"
    Range("AB33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C28"
    Range("AF33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C32"
    Range("T34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1331C20"
    Range("AB34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C28"
    Range("AF34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C32"
    Range("T35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1394C20"
    Range("AB35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C28"
    Range("AF35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C32"
    Range("T36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1457C20"
    Range("AB36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C28"
    Range("AF36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C32"
    Range("T37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1520C20"
    Range("AB37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C28"
    Range("AF37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C32"
    Range("T38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1583C20"
    Range("AB38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C28"
    Range("AF38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C32"
    Range("T39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1646C20"
    Range("AB39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C28"
    Range("AF39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C32"
    Range("T40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1709C20"
    Range("AB40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C28"
    Range("AF40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C32"
    Range("T41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1772C20"
    Range("AB41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C28"
    Range("AF41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C32"
    Range("T42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1835C20"
    Range("AB42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C28"
    Range("AF42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C32"
    Range("T43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1898C20"
    Range("AB43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C28"
    Range("AF43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C32"
    Range("T44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1961C20"
    Range("AB44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C28"
    Range("AF44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C32"
    Range("T45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2024C20"
    Range("AB45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C28"
    Range("AF45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C32"
    Range("T46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2087C20"
    Range("AB46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C28"
    Range("AF46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C32"
    Range("T47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2150C20"
    Range("AB47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C28"
    Range("AF47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C32"
    Range("T48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2213C20"
    Range("AB48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C28"
    Range("AF48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C32"
    Range("T49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2276C20"
    Range("AB49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C28"
    Range("AF49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C32"
    Range("T50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2339C20"
    Range("AB50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C28"
    Range("AF50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C32"
    Range("T51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2402C20"
    Range("AB51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C28"
    Range("AF51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C32"
    Range("T52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2465C20"
    Range("AB52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C28"
    Range("AF52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C32"
    Range("T53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2528C20"
    Range("AB53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C28"
    Range("AF53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C32"
    Range("T54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2591C20"
    Range("AB54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C28"
    Range("AF54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C32"
    Range("T55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2654C20"
    Range("AB55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C28"
    Range("AF55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C32"
    Range("T56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2717C20"
    Range("AB56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C28"
    Range("AF56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C32"
    Range("T57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2780C20"
    Range("AB57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C28"
    Range("AF57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C32"
    Range("T58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2843C20"
    Range("AB58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C28"
    Range("AF58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C32"
    
    
End Sub


Sub Transfert_Mensuel_Décembre_C()


    Range("C13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R8C3"
    Range("K13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C11"
    Range("O13").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R62C15"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R71C3"
    Range("K14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C11"
    Range("O14").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R125C15"
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R134C3"
    Range("K15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C11"
    Range("O15").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R188C15"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R197C3"
    Range("K16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C11"
    Range("O16").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R251C15"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R260C3"
    Range("K17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C11"
    Range("O17").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R314C15"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R323C3"
    Range("K18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C11"
    Range("O18").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R377C15"
    Range("C19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R386C3"
    Range("K19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C11"
    Range("O19").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R440C15"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R449C3"
    Range("K20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C11"
    Range("O20").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R503C15"
    Range("C21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R512C3"
    Range("K21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C11"
    Range("O21").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R566C15"
    Range("C22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R575C3"
    Range("K22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C11"
    Range("O22").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R629C15"
    Range("C23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R638C3"
    Range("K23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C11"
    Range("O23").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R692C15"
    Range("C24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R701C3"
    Range("K24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C11"
    Range("O24").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R755C15"
    Range("C25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R764C3"
    Range("K25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C11"
    Range("O25").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R818C15"
    Range("C26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R827C3"
    Range("K26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C11"
    Range("O26").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R881C15"
    Range("C27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R890C3"
    Range("K27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C11"
    Range("O27").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R944C15"
    Range("C28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R953C3"
    Range("K28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C11"
    Range("O28").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1007C15"
    Range("C29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1016C3"
    Range("K29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C11"
    Range("O29").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1070C15"
    Range("C30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1079C3"
    Range("K30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C11"
    Range("O30").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1133C15"
    Range("C31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1142C3"
    Range("K31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C11"
    Range("O31").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1196C15"
    Range("C32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1205C3"
    Range("K32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C11"
    Range("O32").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1259C15"
    Range("C33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1268C3"
    Range("K33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C11"
    Range("O33").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1322C15"
    Range("C34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1331C3"
    Range("K34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C11"
    Range("O34").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1385C15"
    Range("C35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1394C3"
    Range("K35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C11"
    Range("O35").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1448C15"
    Range("C36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1457C3"
    Range("K36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C11"
    Range("O36").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1511C15"
    Range("C37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1520C3"
    Range("K37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C11"
    Range("O37").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1574C15"
    Range("C38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1583C3"
    Range("K38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C11"
    Range("O38").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1637C15"
    Range("C39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1646C3"
    Range("K39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C11"
    Range("O39").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1700C15"
    Range("C40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1709C3"
    Range("K40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C11"
    Range("O40").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1763C15"
    Range("C41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1772C3"
    Range("K41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C11"
    Range("O41").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1826C15"
    Range("C42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1835C3"
    Range("K42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C11"
    Range("O42").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1889C15"
    Range("C43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1898C3"
    Range("K43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C11"
    Range("O43").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1952C15"
    Range("C44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R1961C3"
    Range("K44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C11"
    Range("O44").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2015C15"
    Range("C45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2024C3"
    Range("K45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C11"
    Range("O45").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2078C15"
    Range("C46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2087C3"
    Range("K46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C11"
    Range("O46").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2141C15"
    Range("C47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2150C3"
    Range("K47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C11"
    Range("O47").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2204C15"
    Range("C48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2213C3"
    Range("K48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C11"
    Range("O48").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2267C15"
    Range("C49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2276C3"
    Range("K49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C11"
    Range("O49").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2330C15"
    Range("C50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2339C3"
    Range("K50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C11"
    Range("O50").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2393C15"
    Range("C51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2402C3"
    Range("K51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C11"
    Range("O51").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2456C15"
    Range("C52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2465C3"
    Range("K52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C11"
    Range("O52").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2519C15"
    Range("C53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2528C3"
    Range("K53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C11"
    Range("O53").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2582C15"
    Range("C54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2591C3"
    Range("K54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C11"
    Range("O54").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2645C15"
    Range("C55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2654C3"
    Range("K55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C11"
    Range("O55").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2708C15"
    Range("C56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2717C3"
    Range("K56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C11"
    Range("O56").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2771C15"
    Range("C57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2780C3"
    Range("K57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C11"
    Range("O57").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2834C15"
    Range("C58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2843C3"
    Range("K58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C11"
    Range("O58").Select
    ActiveCell.FormulaR1C1 = "=[Comptabilité.xls]C!R2897C15"


End Sub


