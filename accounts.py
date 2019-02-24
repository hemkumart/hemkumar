'''
Created on 24-Feb-2019

@author: hem
'''

import xlrd
import xlwt
from xlwt import Workbook
account_workbook=Workbook()
sheet=account_workbook.add_sheet("accounts sheet")

class Account:
    def __init__(self):
        print("Welcome to accounts programing")
    def balance_sheet(self):
        print("balance sheet")
        credit="enter the amount to be credited :"
        debit="enter the amount to debited :"
        balance=credit-debit
        
        sheet.write(1,0,'balance sheet')
        sheet.write(1,1,'credit')
        sheet.write(1,2,'debit')
        sheet.write(1,3,'balance')
        sheet.write(2,1,credit)
        sheet.write(2,2,debit)
        sheet.write(2,3,balance)
        account_workbook.save()
    def asset(self):
        print("asset")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(3,0,asset)     
        sheet.write(3,1,credit)
        sheet.write(3,2,debit)
        sheet.write(3,3,balance)
        account_workbook.save()  
      
    def current(self):
        print("current")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(4,0,current)     
        sheet.write(4,1,credit)
        sheet.write(4,2,debit)
        sheet.write(4,3,balance)
        account_workbook.save()  
    def sundry_debtors(self):
        print("sundry debtors")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(5,0,sundry_debtors) 
        sheet.write(5,1,credit)
        sheet.write(5,2,debit)
        sheet.write(5,3,balance)
        account_workbook.save()
    def accounts_recievable(self):
        print("accounts recievable")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(6,0,accounts_recievable) 
        sheet.write(6,1,credit)
        sheet.write(6,2,debit)
        sheet.write(6,3,balance)
        account_workbook.save()    
    def cash_and_bank(self):
        print("cash_and_bank")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(7,0,cash_and_bank) 
        sheet.write(7,1,credit)
        sheet.write(7,2,debit)
        sheet.write(7,3,balance)
        account_workbook.save()       
    def cash_inhand(self):
        print("cash_inhand")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(8,0,cash_inhand) 
        sheet.write(8,1,credit)
        sheet.write(8,2,debit)
        sheet.write(8,3,balance)
        account_workbook.save()           
        
    def HDFC_BANK_ACCOUNT(self):
        print("HDFC bank account")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(9,0,HDFC_BANK_ACCOUNT) 
        sheet.write(9,1,credit)
        sheet.write(9,2,debit)
        sheet.write(9,3,balance)
        account_workbook.save()           
    def Noncurrent(self):
        print("non current")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(10,0,noncurrent) 
        sheet.write(10,1,credit)
        sheet.write(10,2,debit)
        sheet.write(10,3,balance)
        account_workbook.save()            
    def Fixed(self):
        print("fixed")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(11,0,fixed) 
        sheet.write(11,1,credit)
        sheet.write(11,2,debit)
        sheet.write(11,3,balance)
        account_workbook.save()    
    def Tangibleasset(self):
        print("Tangibleasset")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(12,0,Tangibleasset) 
        sheet.write(12,1,credit)
        sheet.write(12,2,debit)
        sheet.write(12,3,balance)
        account_workbook.save()    
    def furniture_fixes(self):
        print("furniture_fixes")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(13,0,furniture_fixes) 
        sheet.write(13,1,credit)
        sheet.write(13,2,debit)
        sheet.write(13,3,balance)
        account_workbook.save()
        
    def service_equipment(self):
        print("service_equipment")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(14,0,service_equipment) 
        sheet.write(14,1,credit)
        sheet.write(14,2,debit)
        sheet.write(14,3,balance)
        account_workbook.save()
    
    def liablity(self):
        print("liablity")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(15,0,liablity) 
        sheet.write(15,1,credit)
        sheet.write(15,2,debit)
        sheet.write(15,3,balance)
        account_workbook.save()         
    def share_holder_fund(self):
        print("share_holder_fund")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(16,0,share_holder_fund) 
        sheet.write(16,1,credit)
        sheet.write(16,2,debit)
        sheet.write(16,3,balance)
        account_workbook.save()  
    def share_capital_fund(self):
        print("share_capital_fund")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(17,0,share_capita_fund) 
        sheet.write(17,1,credit)
        sheet.write(17,2,debit)
        sheet.write(17,3,balance)
        account_workbook.save()  
    
    def paidup_capital(self):
        print("paidup_capital")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(18,0,paidup_capital) 
        sheet.write(18,1,credit)
        sheet.write(18,2,debit)
        sheet.write(18,3,balance)
        account_workbook.save()  
    
    def mrgrey_capital_AC(self):
        print(" mrgrey_capital_AC")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(19,0, mrgrey_capital_AC) 
        sheet.write(19,1,credit)
        sheet.write(19,2,debit)
        sheet.write(19,3,balance)
        account_workbook.save()      
    def reserve_and_surplus(self):
        print(" reserve_and_surplus")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(20,0, reserve_and_surplus) 
        sheet.write(20,1,credit)
        sheet.write(20,2,debit)
        sheet.write(20,3,balance)
        account_workbook.save()
    def general_reserve(self):
        print(" general_reserve")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(21,0, general_reserve) 
        sheet.write(21,1,credit)
        sheet.write(21,2,debit)
        sheet.write(21,3,balance)
        account_workbook.save()
    def current_liablities(self):
        print(" current_liablities")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(22,0, current_liablities) 
        sheet.write(22,1,credit)
        sheet.write(22,2,debit)
        sheet.write(22,3,balance)
        account_workbook.save()
    def sundry_creditors(self):
        print(" sundry_creditors")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(23,0, sundry_creditors) 
        sheet.write(23,1,credit)
        sheet.write(23,2,debit)
        sheet.write(23,3,balance)
        account_workbook.save()
    def accounts_payable(self):
        print(" accounts_payable")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(24,0, accounts_payable) 
        sheet.write(24,1,credit)
        sheet.write(24,2,debit)
        sheet.write(24,3,balance)
        account_workbook.save()        
    def short_term_loans(self):
        print(" short_term_loans")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(25,0, short_term_loans) 
        sheet.write(25,1,credit)
        sheet.write(25,2,debit)
        sheet.write(25,3,balance)
        account_workbook.save()      
    def loans_payable(self):
        print(" loans_payable")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(26,0, loans_payable) 
        sheet.write(26,1,credit)
        sheet.write(26,2,debit)
        sheet.write(26,3,balance)
        account_workbook.save()     
    def other_current_liablities(self):
        print(" other_current_liablities")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(27,0,  other_current_liablities) 
        sheet.write(27,1,credit)
        sheet.write(27,2,debit)
        sheet.write(27,3,balance)
        account_workbook.save() 
        
    def tax_payable(self):
        print(" tax_payable")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(28,0,tax_payable) 
        sheet.write(28,1,credit)
        sheet.write(28,2,debit)
        sheet.write(28,3,balance)
        account_workbook.save()     
    def gst_payable(self):
        print(" gst_payable")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(29,0,gst_payable) 
        sheet.write(29,1,credit)
        sheet.write(29,2,debit)
        sheet.write(29,3,balance)
        account_workbook.save()     
    def taxes_and_licenses(self):
        print(" taxes_and_licenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(30,0,taxes_and_licenses) 
        sheet.write(30,1,credit)
        sheet.write(30,2,debit)
        sheet.write(30,3,balance)
        account_workbook.save()       
    def profit_and_loss(self):
        print(" profit_and_loss")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(31,0,profit_and_loss) 
        sheet.write(31,1,credit)
        sheet.write(31,2,debit)
        sheet.write(31,3,balance)
        account_workbook.save()
        
    def income(self):
        print(" income")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(32,0,income) 
        sheet.write(32,1,credit)
        sheet.write(32,2,debit)
        sheet.write(32,3,balance)
        account_workbook.save()    
    def directincome(self):
        print(" directincome")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(33,0,directincome) 
        sheet.write(33,1,credit)
        sheet.write(33,2,debit)
        sheet.write(33,3,balance)
        account_workbook.save()       
    def service_revenue(self):
        print(" service_revenue")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(34,0,service_revenue) 
        sheet.write(34,1,credit)
        sheet.write(34,2,debit)
        sheet.write(34,3,balance)
        account_workbook.save()
    def indirectincome(self):
        print(" indirectincome")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(35,0,indirectincome) 
        sheet.write(35,1,credit)
        sheet.write(35,2,debit)
        sheet.write(35,3,balance)
        account_workbook.save()      
    def intrestincome(self):
        print(" intrestincome")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(36,0,intrestincome) 
        sheet.write(36,1,credit)
        sheet.write(36,2,debit)
        sheet.write(36,3,balance)
        account_workbook.save()
    def expenses(self):
        print(" expenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(37,0,expenses) 
        sheet.write(37,1,credit)
        sheet.write(37,2,debit)
        sheet.write(37,3,balance)
        account_workbook.save()    
    def direct_expenses(self):
        print(" direct_expenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(38,0,direct_expenses) 
        sheet.write(38,1,credit)
        sheet.write(38,2,debit)
        sheet.write(38,3,balance)
        account_workbook.save() 
        
        
    def salary_expenses(self):
        print(" salary_expenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(39,0,salary_expenses) 
        sheet.write(39,1,credit)
        sheet.write(39,2,debit)
        sheet.write(39,3,balance)
        account_workbook.save()  
        
    def service_suppliers(self):
        print(" service_suppliers")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(40,0,service_suppliers) 
        sheet.write(40,1,credit)
        sheet.write(40,2,debit)
        sheet.write(40,3,balance)
        account_workbook.save()     
    def indirect_expenses(self):
        print(" indirect_expenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(41,0,indirect_expenses) 
        sheet.write(41,1,credit)
        sheet.write(41,2,debit)
        sheet.write(41,3,balance)
        account_workbook.save()      
    def rental_expenses(self):
        print(" rental_expenses")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(42,0,rental_expenses) 
        sheet.write(42,1,credit)
        sheet.write(42,2,debit)
        sheet.write(42,3,balance)
        account_workbook.save() 
        
    def mrgrey_drawings(self):
        print(" mrgrey_drawings")
        credit=input("enter the amount to be credited :")
        debit=input("enter the amount to debited :")
        balance=credit-debit
        sheet.write(43,0,mrgrey_drawings) 
        sheet.write(43,1,credit)
        sheet.write(43,2,debit)
        sheet.write(43,3,balance)
        account_workbook.save()   
        
    def total(self): 
        my_list1=[]
        my_list2=[] 
        my_list2=[]
        for colns in range(1,sheet.max_colns):
            my_list1.append(colns)
            totalcredit=Sum(my_list1)
            
        for colns in range(2,sheet.max_colns):
            my_list2.append(colns)
            totaldebit=Sum(my_list2)  
        for colns in range(3,sheet.max_colns):
            my_list3.append(colns)
            totalbalance=Sum(my_list3)
    sheet.write(44,1,totalcredit)     
    sheet.write(44,2,totaldebit)
    sheet.write(44,3,totalbalance)
                         
s=Account()
s.balance_sheet()
s.asset()
s.current()
s.sundry_debtors()
s.accounts_recievable()        
s.cash_and_bank()
s.cash_inhand()
s.HDFC_BANK_ACCOUNT()
s.Noncurrent()
s.Fixed()
s.Tangibleasset()
s.furniture_fixes()
s.service_equipment()
s.liablity()
s.share_holder_fund()
s.share_capital_fund()
s.paidup_capital()
s.mrgrey_capital_AC()
s.reserve_and_surplus()
s.general_reserve()
s.current_liablities()
s.sundry_creditors()
s.accounts_payable()
s.short_term_loans()
s.loans_payable()
s.other_current_liablities()
s.tax_payable()
s.gst_payable()
s.taxes_and_licenses()
s.profit_and_loss()
s.income()
s.directincome()
s.service_revenue()
s.indirectincome()
s.intrestincome()
s.expenses()
s.direct_expenses()
s.salary_expenses()
s.service_suppliers()
s.indirect_expenses()
s.rental_expenses()
s.mrgrey_drawings()
s.total()