#!C:\Python27
# -*- coding: utf-8 -*-


class Context():
    """Provides context for Stats Book 1. Variables in context are mainly
    fields of parsed database. Last ones are for internal use of parsers."""

    def __init__(self):

       # title
        self.id_title = None
        self.desc_title = None

        # first level subtitle
        self.id_subt1 = None
        self.desc_subt1 = None

        # second level subtitle
        self.id_subt2 = None
        self.desc_subt2 = None

        # head product table
        self.id_product = None
        self.tariff_number = None
        self.desc_product = None
        self.product_units = None

        # data inside product table
        self.id_country = None
        self.desc_country = None
        self.year = []
        self.quantity = []
        self.value = []

        # context control variables for parsers
        self.row_type = ""
        self.type_last_row = ""
        self.last_row = ""


class RecordsBuilder():
    """Build records from a StatsBook1Context instance."""

    def __init__(self, context):
        self.context = context

    def build_records(self):
        """Build records using the context and modifies it."""

        new_records = []

        # only if row_type is one with data, records can be built
        if self.context.row_type == "agg_values" or \
                self.context.row_type == "tbl_row":

            # each value has a year and quantity, the other variables are equal
            i = 0
            for value in self.context.value:

                new_record = dict()

                # title
                new_record["id_title"] = self.context.id_title
                new_record["desc_title"] = self.context.desc_title

                # first level subtitle
                new_record["id_subt1"] = self.context.id_subt1
                new_record["desc_subt1"] = self.context.desc_subt1

                # second level subtitle
                new_record["id_subt2"] = self.context.id_subt2
                new_record["desc_subt2"] = self.context.desc_subt2

                # head product table
                new_record["id_product"] = self.context.id_product
                new_record["tariff_number"] = self.context.tariff_number
                new_record["desc_product"] = self.context.desc_product
                new_record["product_units"] = self.context.product_units

                # data inside product table
                new_record["id_country"] = self.context.id_country
                new_record["desc_country"] = self.context.desc_country
                new_record["year"] = self.context.year[i]
                new_record["quantity"] = self.context.quantity[i]
                new_record["value"] = value

                # adds new record
                new_records.append(new_record)
                i += 1

        # update last row_type
        self.context.row_type = str(self.context.row_type)

        return new_records

