# sugar2xls

Helper tool dedicated to the retrieval of Sugar CRM data for reporting purpose

## Installation

Pour installer `sugar2xls`utilisez la commande suivante :

    python setup.py install

## Dépendances

`sugar2xls` possède deux dépendances :

 * `xlwt` pour la génération du fichier excel de sortie
 * `python_webservices_library`pour la connexion avec SugarCRM

## Configuration

`sugar2xls` utilise un fichier de configuration JSON pour stocker les informations de base nécessaire pour accéder à SugarCRM. Voici un exemple :

    {
        "url" : "http://crm-test.sysfera.com/sugar/service/v4_1/rest.php",
        "username" : "john.doe",
        "password" : "maxiP4ss"
    }

Ce fichier est nommé `.sugar2xls.config` et il doit se trouver à la racine du compte utilisateur.

## Exécution

`sugar2xls` possède une option permettant de définir un chemin et un nom différent du chemin et du nom par défaut (par défaut le fichier est généré dans le répertoire courant et se nomme data.xls).

Pour spécifier un chemin particulier il faut utiliser l'option `-o` ou `--output`

    sugar2xls.py -o /chemin/vers/mon/fichier/de/sortie/data.xls

## LICENCE

Copyright (C) 2014 David Loureiro

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see [<http://www.gnu.org/licenses/>](http://www.gnu.org/licenses/).