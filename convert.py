#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from lxml import etree
from lxml.etree import *
from lib import xlsx_to_csv as xtc
import os
from datetime import datetime
import unicodecsv

def ut(text):
    return text.decode('utf8')

def getD(row, column):
    # Get the data from the relevant column
    return row[ut(column)]

def get_files_in_dir(relative_dir):
    filenames = []

    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    thedir = os.path.join(current_script_dir, os.path.abspath(relative_dir))

    for file in os.listdir(thedir):
        if file.endswith(".xlsx"):
            abs_filename = os.path.join(thedir, file)
            filenames.append({'absolute_filename': abs_filename,
                              'filename': file,
                              'name': os.path.splitext(file)[0]})
    return filenames

def getCountriesRegions():
    csvfile = open('lib/countries-regions.csv', 'r')
    csv = unicodecsv.DictReader(csvfile)
    cr = {}
    for row in csv:
        cr[row['dac_country_code']] = row
    return cr

def newE(element_name, activity):
    element = Element(element_name)
    activity.append(element)
    return element

def makeISO(date):
    return datetime.strptime(date, "%d/%m/%Y").date().isoformat()

def make_identifier(activity, row, i):
    iati_identifier = newE("iati-identifier", activity)
    extending_org = getD(row, "Agence exécutive")
    project_id = getD(row, "N° de projet du donneur")
    if (project_id == ""):
        project_id = getD(row, "Rubrique MAE") + "-"+str(i)
    iati_identifier.text = "FR-"+extending_org+"-"+project_id
    return iati_identifier

def write_project(doc, row, cr, i):
    getD(row, "Rubrique MAE")

    activity = Element("iati-activity")
    activity.set("default-currency", "EUR")
    nowiso = datetime.now().replace(microsecond=0).isoformat()
    activity.set("last-updated-datetime", nowiso)
    doc.append(activity)

    title = newE("title", activity)
    title.text = getD(row, "Description succinte / titre du projet")

    description = newE("description", activity)
    description.text = getD(row, "Description")

    funding_org = newE("participating-org", activity)
    funding_org.set("role", "Funding")
    funding_org.set("ref", "FR")
    funding_org.text = "France"

    extending_org = newE("participating-org", activity)
    extending_org.set("role", "Extending")
    extending_org.set("ref", "FR-"+getD(row, "Agence exécutive"))
    extending_org.text = getD(row, "Agence exécutive (nom)")

    implementing_org = newE("participating-org", activity)
    implementing_org.set("role", "Implementing")
    implementing_org.set("ref", getD(row, "Canal d'acheminement (code)"))
    implementing_org.text = getD(row, "canal")

    iati_identifier = make_identifier(activity, row, i)

    collab_type = newE("collaboration-type", activity)
    collab_type.set("code", getD(row, "Bi/Multi"))
    collab_type.text = getD(row, """Bi/Multi 
(nom)""")

    flow_type = newE("default-flow-type", activity)
    flow_type.set("code", getD(row, "Type de ressource"))
    flow_type.text = getD(row, "Type de ressource (nom)")

    finance_type = newE("default-finance-type", activity)
    finance_type.set("code", getD(row, "Type de financement"))
    finance_type.text = getD(row, """Type de financement
(nom)""")

    aid_type = newE("default-aid-type", activity)
    aid_type.set("code", getD(row, "Type d'aide"))
    aid_type.text = getD(row, """Type d'aide
(nom)""")

    tied_status = newE("default-tied-status", activity)
    tied_status.set("code", "5")
    tied_status.text = ut("Non lié")

    sector = newE("sector", activity)
    sector.set("code", getD(row, "Secteur (code)"))
    sector.text = getD(row, """Secteur
(nom)""")

    country_region = getD(row, """Pays bénéficiaire du CAD
(code)""")
    
    tcr = cr.get(country_region)
    if (tcr and tcr['type'] =='Country'):
        country = newE("recipient-country", activity)
        country.set('code', tcr['iso2'])
        country.text = tcr['country_name']

    elif (tcr and tcr['type'] == 'Region'):
        region = newE("recipient-region", activity)
        region.set('code', tcr['dac_region_code'])
        region.text = tcr['dac_region_name']

    if getD(row, "Montant versé en millliers  d'euros") != "":
        d_tr = newE("transaction", activity)
        d_tr_type = newE("transaction-type", d_tr)
        d_tr_type.set("code", "D")
        d_tr_type.text = "Versement"
        d_tr_date = newE("transaction-date", d_tr)
        d_tr_date.set("iso-date", "2013-12-31")
        d_tr_value = newE("value", d_tr)
        d_tr_value.set("currency", "EUR")
        d_tr_value.set("value-date", "2013-12-31")
        d_tr_value.text = str(float(getD(row, "Montant versé en millliers  d'euros"))*1000)

    if getD(row, """Montant de l'engagement
en milliers d'euros""") != "":
        c_tr = newE("transaction", activity)
        c_tr_type = newE("transaction-type", c_tr)
        c_tr_type.set("code", "C")
        c_tr_type.text = "Engagement"
        c_tr_date = newE("transaction-date", c_tr)
        c_tr_date.set("iso-date", makeISO(getD(row, "Date d'engagement")))
        c_tr_value = newE("value", c_tr)
        c_tr_value.set("currency", "EUR")
        c_tr_value.set("value-date", "2013-12-31")
        c_tr_value.text = str(float(getD(row, """Montant de l'engagement
en milliers d'euros"""))*1000)

    no = getD(row, "Nature de l'opération (nom)")


def convert():
    cr = getCountriesRegions()
    print get_files_in_dir('source/')

    for XLSX_FILE in get_files_in_dir('source/'):
        doc = Element('iati-activities')
        doc.set("version", "1.05")
        current_datetime = datetime.now().replace(microsecond=0).isoformat()
        doc.set("generated-datetime",current_datetime)

        input_filename = XLSX_FILE['absolute_filename']
        print input_filename
        input_data = open(input_filename).read()
        sheet = 0
        data = xtc.getDataFromFile(input_filename, input_data, sheet, True)
        for i, row in enumerate(data):
            write_project(doc, row, cr, i)
	
        print "Generated activities"
        print "Writing activities ... (4/4)"

        doc = ElementTree(doc)
        doc.write(XLSX_FILE['name']+".xml",encoding='utf-8', xml_declaration=True, pretty_print=True)

if __name__ == '__main__':
    convert()
