# sf-crm-utils

Outils dédié à la gestion de données en lien avec la CRM de SysFera

## Installation

Pour installer les `sf-crm-utils` utilisez la commande suivante :

    python setup.py install

## Dépendances

`sf-crm-utils` possède deux dépendances :

 * `xlwt` pour la génération du fichier excel de sortie
 * `python_webservices_library`pour la connexion avec la CRM

## Configuration

Les `sf-crm-utils` utilisent un fichier de configuration JSON pour stocker les informations de base nécessaire pour accéder à la CRM. Voici un exemple :

    {
        "url" : "http://crm-test.sysfera.com/sugar/service/v4_1/rest.php",
        "username" : "john.doe",
        "password" : "maxiP4ss"
    }

Ce fichier est nommé `.sf-crm.config` et il doit se trouver à la racine du compte utilisateur.

## Exécution

### `sf-crm-getOpportunities`

`sf-crm-getOpportunities` possède une option permettant de définir un chemin et un nom différent du chemin et du nom par défaut (par défaut le fichier est généré dans le répertoire courant et se nomme data.xls).

Pour spécifier un chemin particulier il faut utiliser l'option `-o` ou `--output`

    sf-crm-getOpportunities.py -o /chemin/vers/mon/fichier/de/sortie/data.xls

Le fichier généré va contenir deux onglets:
 
 * Pipe global qui contient toutes les informations des opportunités
 * Factures qui contient l'ensemble des factures à émettre

### `sf-crm-addDocument`

`sf-crm-addDocument` va permettre l'ajout d'un document à la CRM et pour cela il est nécessaire de respecter un pattern : "/.../CODE_OF/Fichier.pdf" avec CODE_OF, l'of du projet considéré.

Il est nécessaire de passer le fichier qui va être ajouté à la CRM en paramètre.

    sf-crm-addDocument.py /chemin/vers/svn/00-Commercial/00-Offres/Client/OF/fichier.pdf

Attention, le code OF doit être précisé dans l'opportunité au sein de la CRM.

## LICENCE

Copyright (C) 2014 David Loureiro

These programs are free software: you can redistribute them and/or modify them under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

These programs are distributed in the hope that they will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with these programs.  If not, see [<http://www.gnu.org/licenses/>](http://www.gnu.org/licenses/).