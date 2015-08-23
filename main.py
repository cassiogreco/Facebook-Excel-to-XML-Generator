#!/usr/bin/python

import xlutil, sys
from lxml import etree
import re
import unicodedata

class XML:

	NUMBER_REQUIRED_COLUMNS = 9

	REQUIRED_COLUMNS = {
		'availability': 'availability',
		'brand': 'brand',
		'condition': 'condition',
		'description': 'description',
		'gtin': 'gtin',
		'id': 'id',
		'image_link': 'image_link',
		'link': 'link',
		'mpn': 'mpn',
		'price': 'price',
		'title': 'title',
	}

	OPTIONAL_COLUMNS = {
		'additional_image_link': 'additional_image_link',
		'age_group': 'age_group',
		'color': 'color',
		'expiration_date': 'expiration_date',
		'gender': 'gender',
		'google_product_category': 'google_product_category',
		'item_group_id': 'item_group_id',
		'material': 'material',
		'pattern': 'pattern',
		'product_type': 'product_type',
		'sale_price': 'sale_price',
		'sale_price_effective_date': 'sale_price_effective_date',
		'shipping': 'shipping',
		'shipping_weight': 'shipping_weight',
		'shipping_size': 'shipping_size',
	}

	VALID_AVAILABILITY_VALUES = {
		'in stock': 'in stock',
		'out of stock': 'out of stock',
	}

	VALID_CONDITION_VALUES = {
		'new': 'new',
		'refurbished': 'refurbished',
		'used': 'used',
	}

	GOOGLE_NAMESPACE_URL = 'http://base.google.com/ns/1.0'

	GOOGLE_NS_ID = 'g'

	VERSION = '2.0'

	XML_TYPE = 'rss'

	ITEM_TAG = 'item'

	ENCODING = 'utf-8'

	DOCUMENTATION_LINK = 'https://developers.facebook.com/docs/marketing-api/'\
		'dynamic-product-ads/product-catalog#XML_RSS'

	def __init__(self, filename=None, excel_filename=None):
		self.file = filename
		self.excel_filename = excel_filename
		self.required_columns_counter = 0
		self.not_accepted_fields = []

	def writeXMLFile(self):
		"""
		Main method that has the logic for reading the excel data and creates the
		valid xml file
		"""
		excel_file_reader = FileReader(self.excel_filename)
		excel_file = excel_file_reader.readFile()
		with open(self.file, 'w+') as XML:
			rss = etree.Element(
				self.XML_TYPE,
				version=self.VERSION,
				nsmap={self.GOOGLE_NS_ID: self.GOOGLE_NAMESPACE_URL},
			)
			channel = etree.Element('channel')
			for line in excel_file:
				item = etree.Element(self.ITEM_TAG)
				# The id tag has to be the first tag
				tag = self._createIdTag(
					line[self.REQUIRED_COLUMNS['id']],
					self.REQUIRED_COLUMNS['id'],
				)
				item.append(tag)
				channel.append(item)
				self.required_columns_counter = 0
				for column in line:
					self._checkRequiredColumns(column)
					if column != self.REQUIRED_COLUMNS['id']:
						if self._checkIfValidColumns(column):
							self._checkValidValue(column, line[column])
							tag = self._createTag(column, line[column])
							item.append(tag)
						elif column not in self.not_accepted_fields:
							self.not_accepted_fields.append(str(column))
					channel.append(item)
			rss.append(channel)
			if len(self.not_accepted_fields) == 0:
				if self.required_columns_counter == self.NUMBER_REQUIRED_COLUMNS:
					xml = etree.tostring(
						rss,
						pretty_print=True,
						xml_declaration=True,
						encoding=self.ENCODING,
					)
					XML.write(xml)
					self._returnToUser(reason='success')
				elif self.required_columns_counter > self.NUMBER_REQUIRED_COLUMNS:
					self._returnToUser(reason='too_many_fields')
				elif self.required_columns_counter < self.NUMBER_REQUIRED_COLUMNS:
					self._returnToUser(reason='missing_fields')
			else:
				self._returnToUser(reason='unsupported_fields')

	def _createIdTag(self, value, tag_name=None):
		"""
		Method that creates the id tag
		"""
		if value is None or value == '' or value == ' ':
			self._returnToUser(reason='empty_field')
			sys.exit()
		if tag_name is not self.REQUIRED_COLUMNS['id']:
			tag_name = self.REQUIRED_COLUMNS['id']
		tag = etree.Element(tag_name)
		tag.text = DataFormatting.idFormatting(value)
		return tag

	def _createTag(self, tag_name, value=None):
		"""
		Method that creates a tag
		"""
		tag = etree.Element(tag_name)
		if tag_name == self.REQUIRED_COLUMNS['price']:
			tag.text = DataFormatting.priceFormatting(value)
			if tag.text == False:
				self._returnToUser(reason='wrong_price_value')
				sys.exit()
		else:
			tag.text = value
		return tag

	def _checkIfValidColumns(self, column_name=None):
		"""
		Method that checks if values are not empty and if they are valid
		"""
		if column_name in self.REQUIRED_COLUMNS:
			return True
		elif column_name in self.OPTIONAL_COLUMNS:
			return True
		return False

	def _checkValidValue(self, column=None, value=None):
		"""
		Method that checks if values are not empty and if they are valid
		"""
		if column is None or column == '' or column == ' ' or \
			value is None or value == '' or value == ' ':
			self._returnToUser(reason='missing_fields')
		if column == 'availability':
			if value not in self.VALID_AVAILABILITY_VALUES:
				self._returnToUser(
					reason='wrong_availability_value',
				 	extra_value=value,
				 )
		if column == 'condition':
			if value not in self.VALID_CONDITION_VALUES:
				self._returnToUser(
					reason='wrong_condition_value',
					extra_value=value,
				)
				
	def _checkRequiredColumns(self, column_name=None):
		"""
		Method that checks if the column is one of the required columns. If so,
		increment the counter
		"""
		if column_name in self.REQUIRED_COLUMNS:
			self.required_columns_counter += 1

	def _returnToUser(self, reason=None, extra_value=None):
		"""
		Method that returns a success or an appropriate error message
		"""
		if reason == 'success':
			print('\nXML File created!\n')
		if reason == 'too_many_fields':
			print('\nError: You have too many required fields. There should '\
				'only be 9. Please check the documentation at: {}\n'
				.format(self.DOCUMENTATION_LINK))
		if reason == 'missing_fields':
			print('\nError: {} required fields are missing. They might have '\
				'been misspelled Please check the documentation at: {}\n'
				.format(
					self.required_columns_counter,
					self.DOCUMENTATION_LINK,
				))
		if reason == 'unsupported_fields':
			print('\nError: There are fields in the excel that we do not '\
				'support. They are: {}. Please check the documentation at: {}\n'
				.format(
					self.not_accepted_fields,
					self.DOCUMENTATION_LINK,
				))
		if reason == 'empty_field':
			print('\nError: There is at least one empty field in the excel. '\
				'We do not accept empty tags. Please check the documentation '\
				'at: {}\n'.format(
					self.DOCUMENTATION_LINK,
				))
		if reason == 'wrong_availability_value':
			print('\nError: Availability column does not have the supported'\
				' value. Value sent = {}. Please check the documentation '\
				'at: {}\n'.format(
					extra_value,
					self.DOCUMENTATION_LINK,
				))
		if reason == 'wrong_condition_value':
			print('\nError: Condition column does not have the supported'\
				' value. Value sent = {}. Please check the documentation '\
				'at: {}\n'.format(
					extra_value,
					self.DOCUMENTATION_LINK,
				))
		if reason == 'wrong_price_value':
			print('\nError: Price column does not have the supported'\
				' value. Value sent = {}. Please check the documentation '\
				'at: {}\n'.format(
					extra_value,
					self.DOCUMENTATION_LINK,
				))

class DataFormatting:

	def __init__(self):
		pass

	@staticmethod
	def priceFormatting(price):
		"""
		Method that checks and correctly formats the price tag. Does not check
		all possible values. For example, alphanumeric values like 123asdf will
		pass, but Facebook's check will get those.
		"""
		formatted_price = None
		if re.search('[0-9]', formatted_price or price):
			if price[len(price) - 3:] != 'BRL':
				price += ' BRL'
			if not re.search('\.', price):
				for i in xrange(0, price):
					if price[i] == ' ':
						formatted_price = price[:i-3] + '.00' + price[i-3:]
			elif re.search(',', formatted_price or price):
				stringed_price = unicodedata.normalize(
					'NFKD',
					price,
				).encode('ascii','ignore')
				periods_removed = stringed_price.replace('.', '')
				formatted_price = periods_removed.replace(',', '.')
			elif re.search(',', formatted_price or price):
				stringed_price = unicodedata.normalize(
					'NFKD',
					price,
				).encode('ascii','ignore')
				formatted_price = stringed_price.replace(',', '')
			if formatted_price is not None:
				return formatted_price
		else:
			return False
		return price

	@staticmethod
	def idFormatting(id):
		"""
		Method that returns the correct id format. Values coming from the excel
		sheet come in as: u'2.0' instead of 2
		"""
		return id[:len(id)-2]


class FileReader:

	def __init__(self, filename):
		self.filename = filename

	def readFile(self):
		"""
		Method that creates the excel file reader
		"""
		return xlutil.one_sheet_with_headers(self.filename)


def main(filename, excel_filename):
	xml = XML(filename, excel_filename)
	xml.writeXMLFile()


if __name__ == "__main__":
	"""
	Start the program with a commend like:
	python main.py new_generated_xml_filename.xml my_excel_file.xlsx
	"""
	main(sys.argv[1], sys.argv[2])
