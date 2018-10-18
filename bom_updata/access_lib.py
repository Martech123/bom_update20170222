#encoding=utf-8
#author: walker
#date: 2015-09-25
#function: 打印指定access文件、指定表的所有字段

import pypyodbc

excel_file = r'D:\workspace\Python\bom_edit\cnu_c8811_v2.xls'
access_file = r'Driver={Microsoft Access Driver (*.mdb)};DBQ=D:\workspace\Python\bom_edit\credo_cis_lib.mdb'




def read_access_quantity(myID,access_file): # myID:str access_file:str 



    """
    IC='104'   # 104 --->IC
    Capacitor='102'   # 102----->Capacitor
    ='102'   # 102----->Capacitor
    Crystals='107'  
    inductor='103'
    Resistor='101'
    """

   # print('################# start function read_quantity() ######################')
    ID_state = 0
    q = 0
    quantity_col = 0

    PP_TYPE=''
    if myID == '':
        # print('ID 输入为空')
        return ID_state,q


##############################################################
##  ID 匹配



    P_TYPE=myID[:3]


    if P_TYPE=='102':
        #print('C')
        PP_TYPE='Capacitor'
    elif P_TYPE=='107':
        #print('Cry')
        PP_TYPE='Crystals'
    elif P_TYPE=='104':
        #print('IC')
        PP_TYPE='IC'
    elif P_TYPE =='113':
        PP_TYPE='Connector'
    elif P_TYPE=='103':
        PP_TYPE='inductor'
    elif P_TYPE =='115':
        PP_TYPE='PD'
    elif P_TYPE=='101':
        PP_TYPE='Resistor'
    elif P_TYPE=='311':
        PP_TYPE='EVB'
    else:
        # print('ID 不存在')
        return ID_state,q
    #print('*******************')
    # print('Type='+PP_TYPE)

#########################################################



    #print('*******************')
    if PP_TYPE!='':
        Select="SELECT * FROM " + PP_TYPE + " WHERE PNID="+"'"+myID+"'"
        print('sql = '+Select)
    else:
        Select=''



    conn = pypyodbc.connect(access_file)
    cur = conn.cursor()
    cur.execute(Select)

    mdblist = cur.fetchone()
    # print(mdblist)


    if mdblist!=None:
        # print(mdblist)
        quantity_col = get_access_field_coord(PP_TYPE,'Qty',access_file)
        # print(quantity_col)
        q = mdblist[quantity_col]
        # print('q=',q)
        ID_state = 1
    else:
        ID_state = 0
        q = 0 


        cur.close()
        conn.close()
        #print('################# End function read_quantity() ######################')
        #print()
        #print()



    return ID_state,q


def update_access_quantity(myID,q,access_file): #myID,q,access_file --->str
    print('################ start function update_quantity() ######################')
 
##############################################################
##  ID 匹配

    ID_state = myID
    q = q
   
    PP_TYPE=''
    if myID == '':
        # print('ID 输入为空')
        return ID_state,q


    P_TYPE=myID[:3]


    if P_TYPE=='102':
        #print('C')
        PP_TYPE='Capacitor'
    elif P_TYPE=='107':
        #print('Cry')
        PP_TYPE='Crystals'
    elif P_TYPE=='104':
        #print('IC')
        PP_TYPE='IC'
    elif P_TYPE =='113':
        PP_TYPE='Connector'
    elif P_TYPE=='103':
        PP_TYPE='inductor'
    elif P_TYPE =='115':
        PP_TYPE='PD'
    elif P_TYPE=='101':
        PP_TYPE='Resistor'
    else:
        # print('ID 不存在')
        return ID_state,q
    #print('*******************')
    # print('Type='+PP_TYPE)

#########################################################
    # P_TYPE=myID[:3]

    # PP_TYPE=''
    # if P_TYPE=='102':
    #     #print('C')
    #     PP_TYPE='Capacitor'
    # elif P_TYPE=='107':
    #     #print('Cry')
    #     PP_TYPE='Crystals'
    # elif P_TYPE=='104':
    #     #print('IC')
    #     PP_TYPE='IC'
    # else:
    #     PP_TYPE=''
    # print('*******************')
    # print('Type='+PP_TYPE)






    print('*******************')
    if PP_TYPE!='':
        update_comm="UPDATE " + PP_TYPE +" SET Qty='"+q+"' WHERE PNID="+"'"+myID+"'"
        print('update_comm='+update_comm)
    else:
        Select=''

    conn = pypyodbc.connect(access_file)
    cur = conn.cursor()
    cur.execute(update_comm)
    print('update successfull')
    cur.close()
    conn.commit()
    conn.close()

    print('################ End function update_quantity() ######################')
    print()
    print()

def get_access_field_coord(p_type,name,access_file):

    conn = pypyodbc.connect(access_file)
    cur = conn.cursor()
    Select = 'SELECT * from '+p_type
    cur.execute(Select)
    mdblist = cur.description
    # print(mdblist)
    # print(mdblist[1][0])

    # print(len(mdblist))



    a=0
    b=0
    for i in range(len(mdblist)):

        # print(mdblist[i][0])
        if mdblist[i][0]==name:
            b = a
        a = a+1
    # print('11111111111111111111111')


    cur.close()
    conn.close()
    return b



if __name__ == '__main__':
    print('hello 123123123123123123123')
    # read_access_quantity('101004',access_file)
    update_access_quantity('101003','9999',access_file)