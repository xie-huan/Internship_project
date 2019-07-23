# encoding:gbk
require_relative 'TypeFactory'

# ������������

# #���Ա���Ľṹ
# TEST_REPORT_ITEM = {
#   #sheet��=>��ͷ
#   :CheckItem => ["Checker","CheckItem"],
#   :SrvDetail => ["Path","Number","DstVer","NeId","Action","Instru","BidResetTypeInProt","SrvCheckElement","SrvStatus","InterruptTime"],
#   :BdverInfo => ["Path","Number","NeId","DstVer","Bid","BdType","SrcSoftVer","DstSoftVer","SrcLogicVer","DstLogicVer","SrcCpldVer","DstCpldVer","SrcPcbVer","DstPcbVer"],
#   :Detail => ["Path","Number","NeId","DstVer",'Action','Checker','Result','ErrInfo','Errlog'],
#   :Overview => ["Path","Number",'NeId',"DstVer",'Result','TimeUsed'],
# }

#���Ա���Ľṹ
TEST_REPORT_ITEM = {
    #sheet��=>��ͷ
    :CheckItem => ["Checker","CheckItem"],
    :SrvDetail => ["Path","Number","DstVer","NeId","Action","Instru","Bid-ResetType-InProt","SrvCheckElement","SrvStatus","InterruptTime(Longest)"],
    :BdverInfo => ["Path","Number","NeId","DstVer","Bid","BdType","SrcSoftVer","DstSoftVer","SrcLogicVer","DstLogicVer","SrcCpldVer","DstCpldVer","SrcPcbVer","DstPcbVer"],
    :Detail => ["Path","Number","NeId","DstVer",'Action','Checker','Result','ErrInfo','Errlog'],
    :Overview => ["Path","Number",'NeId',"DstVer",'Result','TimeUsed(minutes)'],
}
#
# #------------------excel ����----------------------
excle = TypeFactory.instance.create("Excel",TEST_REPORT_ITEM)
excle.record("Detail","Path","NumberOfTimes","3-75","DstVer",'Action','Checker','Result','ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221121')
excle.record("Detail","Path1","NumberOfTimes1","3-75","DstVer1",'Action1','1Checker','1Result','1ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221115')
excle.record("CheckItem","Path1","NumberOfTimes1")
excle.close


#------------------mysqldb ����---------------------
# begin
#   mysqldb = TypeFactory.instance.create("MysqlDB",TEST_REPORT_ITEM,"root","xiehuan+123456")
#   mysqldb.record("Detail","Path","NumberOfTimes","3-75","DstVer",'Action','Checker','Result','ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221121')
#   mysqldb.record("Detail","Path1","NumberOfTimes1","3-75","DstVer1",'Action1','1Checker','1Result','1ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221115')
#   mysqldb.record("CheckItem","Path1","NumberOfTimes1")
#   # �ر����ݿ�����
#   mysqldb.close
# rescue
#   p "�������ӻ�������ݿ�ʧ��"
# end


#------------------html ����---------------------

html = TypeFactory.instance.create("HTML",TEST_REPORT_ITEM)
html.record("Detail","Path","NumberOfTimes","3-75","DstVer",'Action','Checker','Result','ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221121')
html.record("Detail","Path1","NumberOfTimes1","3-75","DstVer1",'Action1','1Checker','1Result','1ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221115')
html.record("CheckItem","Path1","NumberOfTimes1")
html.close


#------------------mongodb ����---------------------
mongo = TypeFactory.instance.create("MongoDB",TEST_REPORT_ITEM,"test","pwd")
mongo.record("Detail","Path","NumberOfTimes","3-75","DstVer",'Action','Checker','Result','ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221121')
mongo.record("Detail","Path1","NumberOfTimes1","3-75","DstVer1",'Action1','1Checker','1Result','1ErrInfo','D:\UpgradeTester\output\nelog\2-64\20190410221115')
mongo.record("CheckItem","Path1","NumberOfTimes1")
mongo.record("Overview",1,1,1,1,1,1)
mongo.close
