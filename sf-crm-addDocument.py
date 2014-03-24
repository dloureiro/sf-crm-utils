#!/usr/bin/env python
#
# sf-crm-addDocument.py
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
import base64

VERSION = "1.0"

def main():

    parser = OptionParser(usage="usage: sf-crm-addDocument filepath",
                          version="sf-crm-addDocument " + VERSION)
    (options, args) = parser.parse_args()

    document = args[0]

    fileName = os.path.basename(document)
    path = os.path.dirname(document)
    ofCode = os.path.basename(path)

    #print "ofCode " + ofCode
    #print "filename " + fileName

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
    #data = conn.get_module_fields("Revisions")

    #pp.pprint(data)
    #exit()

    print "Recuperation de l'opportunite correspondant a l'OF " + ofCode

    #pp.pprint(aDocument)

    query = conn.modules["Opportunities"].query()
    opp = query.filter(of_code_c__contains=ofCode)
    opportunity = opp[0]

        # This new query has a filter. Please notice that the filter parameter is the
    # field name in the SugarCRM module, followed by a double underscore, and then
    # an operator (it can be 'exact', 'contains', 'gt', 'gte', 'lt', 'lte' or 'in').

    query = conn.modules['Documents'].query()

    print "Creation d'un document"
    aDocument = sugarcrm.SugarEntry(conn.modules["Documents"])
    aDocument["name"] = fileName 
    aDocument["filename"]=fileName 
    aDocument["document_name"]=fileName 
    aDocument.save()

    print aDocument["id"]

    with open(document) as f:
        encoded = base64.b64encode(f.read())
        print "encoded : " + encoded
        conn.set_document_revision({"id": aDocument["id"],
            "revision":1,
            "filename":fileName + "-test",
            "file":encoded})


    print "Liaison entre le doc cree et "+ opportunity["name"]
    aDocument.relate(opportunity)

if __name__ == '__main__':
    main()
