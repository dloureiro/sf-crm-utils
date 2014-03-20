#!/usr/bin/env python
#
# sugar2xls.py
# Copyright (C) <year>  <name of author>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import sugarcrm
import xlwt
import datetime
import os.path
import json
import pprint
from optparse import OptionParser

def main():

    parser = OptionParser(usage="usage: sugar2xls [options]",
                          version="sugar2xls 1.0")
    parser.add_option("-o", "--output",
                      action="store_true",
                      dest="output",
                      default="data.xls",
                      help="Chemin vers le fichier de sortie contenant les donnees recuperees de SugarCRM")
    (options, args) = parser.parse_args()

    #print options
    #print args

    output = options.output

    home = os.path.expanduser("~")

    configFile=os.path.join(home,".sugar2xls.config")

    if not os.path.isfile(configFile) :
        print "Error : config file " + configFile + " does not exist"
        exit()

    config = json.load(open(configFile,"r"))

    pp = pprint.PrettyPrinter(indent=4)

    #pp.pprint(config)

    # This is the URL for the v4 REST API in your SugarCRM server.
    url = config["url"]
    username = config["username"]
    password = config["password"]


    # This way you log-in to your SugarCRM instance.
    conn = sugarcrm.Sugarcrm(url, username, password)
    data = conn.get_module_fields("Opportunities")

    #pp.pprint(data)
    #exit()

    # This new query has a filter. Please notice that the filter parameter is the
    # field name in the SugarCRM module, followed by a double underscore, and then
    # an operator (it can be 'exact', 'contains', 'gt', 'gte', 'lt', 'lte' or 'in').

    query = conn.modules['Opportunities'].query()
    good_opportunities= query.filter(date_closed__gte='2014-01-01',date_closed__lt='2015-01-01',
        opportunity_status_c__exact='active')

    wb = xlwt.Workbook(encoding="latin-1")
    wsPipeGlobal = wb.add_sheet('Pipe global')

    print "Writing head of 'Pipe global'"

    wsPipeGlobal.write(0, 0, u"Nom")
    wsPipeGlobal.write(0, 1, u"Compte")
    wsPipeGlobal.write(0, 2, u"Date de closing")
    wsPipeGlobal.write(0, 3, u"Sales stage")
    wsPipeGlobal.write(0, 4, u"Type d'opportunite")
    wsPipeGlobal.write(0, 5, u"Probabilite")
    wsPipeGlobal.write(0, 6, u"Montant")
    wsPipeGlobal.write(0, 7, u"Manager")
    wsPipeGlobal.write(0, 8, u"Delai facturation")
    wsPipeGlobal.write(0, 9, u"Pourcentage d'acompte")
    wsPipeGlobal.write(0,10, u"Date de fin de projet")
    wsPipeGlobal.write(0,11, u"Date d'entree")
    wsPipeGlobal.write(0,12, u"Date de modification")


    rowIndex=1

    print "Writing Opportunities"

    for opportunity in good_opportunities:

        # Getting opportunity elements
        name = opportunity['name']
        manager = opportunity['assigned_user_name']

        compte = opportunity["account_name"]
        date = opportunity["date_closed"]
        sales_stage = opportunity['sales_stage']
        opportunity_type = opportunity['type_opportunite__c']
        probabilite = opportunity["probability"]
        amount = opportunity["amount"]
        delai_facturation = opportunity["delai_facturation_c"]
        acompte_pourcentage = opportunity["acompte_pourcentage_c"]
        projet_end_date = opportunity["projet_end_date_c"]
        date_entered = opportunity["date_entered"]
        date_modified = opportunity["date_modified"]

        print "Writing opportunity " + name

        # ecriture du nom d'opportunite et du compte associe
        wsPipeGlobal.write(rowIndex,0,name.encode('latin-1')) #A
        wsPipeGlobal.write(rowIndex,1,compte.encode('latin-1')) #B

        # date management
        year,month,day = date.split('-')

        style = xlwt.XFStyle()
        style.num_format_str = "DD/MM/YY"

        wsPipeGlobal.write(rowIndex,2,datetime.date(int(year),int(month),int(day)),style) #C

        # sales stage
        wsPipeGlobal.write(rowIndex,3,sales_stage.encode('latin-1')) #D

        # opportunity type
        wsPipeGlobal.write(rowIndex,4,opportunity_type.encode('latin-1')) #E

        # probabilite
        proba = float(probabilite)/100

        style = xlwt.XFStyle()
        style.num_format_str = "0%"

        wsPipeGlobal.write(rowIndex,5,proba,style) #F

        # montant
        amount = float(amount)

        wsPipeGlobal.write(rowIndex,6,amount) #G

        # manager de l'affaire
        wsPipeGlobal.write(rowIndex,7,manager.encode('latin-1'))

        # delai facturation
        if delai_facturation=='':
            delai_facturation='0'
        wsPipeGlobal.write(rowIndex,8,int(delai_facturation))

        if acompte_pourcentage=='':
            acompte_pourcentage='0'
        pourcentageAcompte = float(acompte_pourcentage)/100

        wsPipeGlobal.write(rowIndex,9,pourcentageAcompte,style)
        
        if projet_end_date=='':
            projet_end_date=year + '-' + str(int(month)+3)+'-'+day

        style = xlwt.XFStyle()
        style.num_format_str = "DD/MM/YY"

        year,month,day = projet_end_date.split('-')

        wsPipeGlobal.write(rowIndex,10,datetime.date(int(year),int(month),int(day)),style) 

        styleLong = xlwt.XFStyle()
        styleLong.num_format_str = "DD/MM/SS hh:mm:ss"
        
        yymmdd,hhmmss = date_entered.split(" ")

        year,month,day = yymmdd.split('-')
        hour,minute,second = hhmmss.split(':')

        wsPipeGlobal.write(rowIndex,11,datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), int(second)),styleLong) 

        yymmdd,hhmmss = date_modified.split(" ")

        year,month,day = yymmdd.split('-')
        hour,minute,second = hhmmss.split(':')

        wsPipeGlobal.write(rowIndex,12,datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), int(second)),styleLong) 

        rowIndex = rowIndex + 1

    print "Outputting to file"
    wb.save(output)

if __name__ == '__main__':
    main()
