#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from lxml import etree
from lxml.etree import *
from lib import xlsx_to_csv as xtc
import os
from datetime import datetime, timedelta
import unicodecsv
import unicodedata, re

def ut(text):
    return text.decode('utf8')

def remove_control_characters(s):
    return "".join(ch for ch in s if unicodedata.category(ch)[0]!="C")

def getD(row, column):
    # Get the data from the relevant column
    y = row[ut(column)]
    if type(y) == str:
        y = y.decode("utf8", "replace")
        y = remove_control_characters(y)
    return y

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

def getCodesCountries():
    csvfile = open('lib/countries-regions.csv', 'r')
    csv = unicodecsv.DictReader(csvfile)
    cr = {}
    for row in csv:
        cr[row['country_name_FR']] = row
    return cr

codesCountries = getCodesCountries()

def getStatus(value):
    if value == "En cours":
        return ("2", u"En cours")
    return ("3", u"Terminé")

def getCodeFromCountry(value):
    country = codesCountries.get(value)
    if not country:
        return "998"
    return country.get("dac_country_code")

def newE(element_name, activity):
    element = Element(element_name)
    activity.append(element)
    return element

def makeISO(date):
    return datetime.strptime(date, "%d/%m/%Y").date().isoformat()

def makeISOYear(date):
    try:
        return datetime.strptime(date, "%Y").date().isoformat()
    except ValueError:
        return "2015-01-01"

def makeISOYearEnd(date):
    try:
        year = str(int(date)+1)
        yeardate = datetime.strptime(year, "%Y").date()
        yeardate = yeardate - timedelta(days=1)
        return yeardate.isoformat()
    except ValueError:
        return "2015-12-31"

def make_identifier(activity, row, i):
    iati_identifier = newE("iati-identifier", activity)
    iati_identifier.text = "FR-DAECT-%s" % i
    return iati_identifier

def write_project(doc, row, cr, i):

    activity = Element("iati-activity")
    activity.set("default-currency", "EUR")
    nowiso = datetime.now().replace(microsecond=0).isoformat()
    activity.set("last-updated-datetime", nowiso)
    doc.append(activity)

    title = newE("title", activity)
    title.text = getD(row, "Intitulé du projet")

    description = newE("description", activity)
    description.text = getD(row, "Description détaillée")

    reporting_org = newE("reporting-org", activity)
    reporting_org.set("ref", "FR-6")
    reporting_org.text = "France"

    funding_org = newE("participating-org", activity)
    funding_org.set("role", "Funding")
    funding_org.set("ref", "FR")
    funding_org.text = "France"

    extending_org = newE("participating-org", activity)
    extending_org.set("role", "Extending")
    extending_org.set("ref", "FR-99")
    extending_org.text = getD(row, "Collectivité")

    implementing_org = newE("participating-org", activity)
    implementing_org.set("role", "Implementing")
    implementing_org.text = getD(row, "Partenaire")

    iati_identifier = make_identifier(activity, row, i)

    collab_type = newE("collaboration-type", activity)
    collab_type.set("code", "1")
    collab_type.text = ut("Bilatéral")

    planned_start = newE("activity-date", activity)
    planned_start.set("type", "start-planned")
    planned_start.set("iso-date", makeISOYear(getD(row, "Année de début")))

    flow_type = newE("default-flow-type", activity)
    flow_type.set("code", "10")
    flow_type.text = ut("APD (aide publique au développement)")

    finance_type = newE("default-finance-type", activity)
    finance_type.set("code", "110")
    finance_type.text = ut("Don sauf réorganisation de la dette")

    aid_type = newE("default-aid-type", activity)
    aid_type.set("code", "C01")
    aid_type.text = ut("Interventions de type projet")

    tied_status = newE("default-tied-status", activity)
    tied_status.set("code", "5")
    tied_status.text = ut("Non lié")

    theStatus = getStatus(getD(row, "Statut"))
    status = newE("activity-status", activity)
    status.set("code", theStatus[0])
    status.text = theStatus[1]

    sector = newE("sector", activity)
    sector.set("vocabulary", "RO")
    sector.text = getD(row, """Thématique""")

    sector = newE("sector", activity)
    sector.set("vocabulary", "DAC")
    sector.set("code", "43010")
    sector.text = "Aide plurisectorielle"

    country_region = getCodeFromCountry(getD(row, "Pays"))
    
    tcr = cr.get(country_region)
    if (tcr and tcr['type'] =='Country'):
        country = newE("recipient-country", activity)
        country.set('code', tcr['iso2'])
        country.text = tcr['country_name']

    elif (tcr and tcr['type'] == 'Region'):
        region = newE("recipient-region", activity)
        region.set('code', tcr['dac_region_code'])
        region.text = tcr['dac_region_name']

    if getD(row, "Coût total de l’opération") != "0":
        d_tr = newE("transaction", activity)
        d_tr_type = newE("transaction-type", d_tr)
        d_tr_type.set("code", "C")
        d_tr_type.text = "Engagement"
        d_tr_date = newE("transaction-date", d_tr)
        d_tr_date.set("iso-date", makeISOYear(getD(row, "Année de début")))
        d_tr_value = newE("value", d_tr)
        d_tr_value.set("currency", "EUR")
        d_tr_value.set("value-date", makeISOYear(getD(row, "Année de début")))
        d_tr_value.text = str(float(getD(row, "Coût total de l’opération")))
        d_tr_description = newE("description", d_tr)
        d_tr_description.text = u"Année de début du projet; montant total"

    if getD(row, "Cofinancement du MAEDI") != u"0":
        d_tr = newE("transaction", activity)
        d_tr_type = newE("transaction-type", d_tr)
        d_tr_type.set("code", "D")
        d_tr_type.text = "Versement"
        d_tr_date = newE("transaction-date", d_tr)
        d_tr_date.set("iso-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value = newE("value", d_tr)
        d_tr_value.set("currency", "EUR")
        d_tr_value.set("value-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value.text = str(float(getD(row, "Cofinancement du MAEDI")))
        d_tr_provider = newE("provider-org", d_tr)
        d_tr_provider.set("ref", "FR-6")
        d_tr_provider.text = "MAEDI"
        d_tr_description = newE("description", d_tr)
        d_tr_description.text = u"Fin de l'année de début du projet; cofinancement du MAEDI"

    if getD(row, "Part financière des collectivités") != u"0":
        d_tr = newE("transaction", activity)
        d_tr_type = newE("transaction-type", d_tr)
        d_tr_type.set("code", "D")
        d_tr_type.text = "Versement"
        d_tr_date = newE("transaction-date", d_tr)
        d_tr_date.set("iso-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value = newE("value", d_tr)
        d_tr_value.set("currency", "EUR")
        d_tr_value.set("value-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value.text = str(float(getD(row, "Part financière des collectivités")))
        d_tr_provider = newE("provider-org", d_tr)
        d_tr_provider.set("ref", "FR-99")
        d_tr_provider.text = getD(row, "Collectivité")
        d_tr_description = newE("description", d_tr)
        d_tr_description.text = u"Fin de l'année de début du projet; cofinancement du collectivité"

    if getD(row, "Total autre financement") != u"0":
        d_tr = newE("transaction", activity)
        d_tr_type = newE("transaction-type", d_tr)
        d_tr_type.set("code", "D")
        d_tr_type.text = "Versement"
        d_tr_date = newE("transaction-date", d_tr)
        d_tr_date.set("iso-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value = newE("value", d_tr)
        d_tr_value.set("currency", "EUR")
        d_tr_value.set("value-date", makeISOYearEnd(getD(row, "Année de début")))
        d_tr_value.text = str(float(getD(row, "Total autre financement")))
        d_tr_provider = newE("provider-org", d_tr)
        d_tr_provider.text = getD(row, "Autre")
        d_tr_description = newE("description", d_tr)
        d_tr_description.text = u"Fin de l'année de début du projet; cofinancement d'autres organisations"

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
        print "Writing activities ..."

        doc = ElementTree(doc)
        doc.write(XLSX_FILE['name']+".xml",encoding='utf-8', xml_declaration=True, pretty_print=True)

if __name__ == '__main__':
    convert()
