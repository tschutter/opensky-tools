#!/usr/bin/env python

"""
Each row in a .csv file exported from OpenSky corresponds to a
single product in an order.  This script merges those rows into a
single order.

The output .csv file has been successfully imported into
gnucash-2.6.1.  It has not been tested with Quicken.

QIF format: http://en.wikipedia.org/wiki/QIF
"""

from __future__ import print_function
import argparse
import csv
import os
import sys

class Order:
    def __init__(
        self,
        date,
        sku,
        item_price,
        shipping,
        opensky_credits,
        sales_tax,
        restocking_fee,
        cc_processing,
        opensky_commission,
        total_payment
    ):
        self.date = date
        self.skus = sku
        self.item_price = item_price
        self.shipping = shipping
        self.opensky_credits = opensky_credits
        self.sales_tax = sales_tax
        self.restocking_fee = restocking_fee
        self.cc_processing = cc_processing
        self.opensky_commission = opensky_commission
        self.total_payment = total_payment

    def update(
        self,
        sku,
        item_price,
        shipping,
        opensky_credits,
        sales_tax,
        restocking_fee,
        cc_processing,
        opensky_commission,
        total_payment
    ):
        self.skus += "," + sku
        self.item_price += item_price
        self.shipping += shipping
        self.opensky_credits += opensky_credits
        self.sales_tax += sales_tax
        self.restocking_fee += restocking_fee
        self.cc_processing += cc_processing
        self.opensky_commission += opensky_commission
        self.total_payment += total_payment

    def validate(self, order_id):
        item_price = round(
            self.total_payment - (
                self.shipping +
                self.opensky_credits +
                self.sales_tax +
                self.restocking_fee +
                self.cc_processing +
                self.opensky_commission
            ),
            2
        )
        if self.item_price != item_price:
            print(
                "WARNING: For Order ID = {}, Item price {:6.2f} != calculated price {:6.2f}".format(
                    order_id,
                    self.item_price,
                    item_price
                ),
                file=sys.stdout
            )
            self.item_price = item_price

def get_float(fields, key):
    """Convert a string value to a float, handling blank values."""
    value = fields[key]
    try:
        value = float(value)
    except ValueError:
        value = 0.0
    return value

def read_opensky(args):
    """Read a .csv file exported from OpenSky."""
    orders = {}
    with open(args.csvfile) as opensky_file:
        opensky_reader = csv.DictReader(opensky_file)
        for rownum, fields in enumerate(opensky_reader):
            order_id = fields["Order ID"]
            sku = fields["SKU"]
            item_price = get_float(fields, "Item price")
            shipping = get_float(fields, "Shipping price")
            # Sign on credits is wrong in exported .csv file.
            opensky_credits = -get_float(fields, "Credits")
            sales_tax = get_float(fields, "Sales tax")
            restocking_fee = get_float(fields, "Restocking fee")
            cc_processing = get_float(fields, "Credit card processing")
            opensky_commission = get_float(fields, "OpenSky commission")
            total_payment = get_float(fields, "Total payment")
            date = fields["Original order date"]
            if order_id in orders:
                orders[order_id].update(
                    sku,
                    item_price,
                    shipping,
                    opensky_credits,
                    sales_tax,
                    restocking_fee,
                    cc_processing,
                    opensky_commission,
                    total_payment
                )
            else:
                orders[order_id] = Order(
                    date,
                    sku,
                    item_price,
                    shipping,
                    opensky_credits,
                    sales_tax,
                    restocking_fee,
                    cc_processing,
                    opensky_commission,
                    total_payment
                )

    return orders

def validate_orders(orders):
    """Check that the numbers add up for each order."""
    for order_id in orders:
        orders[order_id].validate(order_id)

def write_split(qif_file, amount, account, memo=None):
    """Write a single split to a .qif file."""
    if amount != 0.0:
        print("S{}".format(account), file=qif_file)
        if memo:
            print("E{}".format(memo), file=qif_file)
        print("${:.2f}".format(amount), file=qif_file)

def write_qif(args, orders):
    """Write a .qif file."""
    with open(args.qiffile, "w") as qif_file:
        print("!Account", file=qif_file)
        print("N{}".format(args.acct_opensky), file=qif_file)
        print("TBank", file=qif_file)
        print("^", file=qif_file)
        print("!Type:Bank", file=qif_file)
        for order_id in sorted(orders):
            order = orders[order_id]
            print("D{}".format(order.date), file=qif_file)
            print("POpenSky order {}".format(order_id), file=qif_file)
            print("L{}".format(args.acct_opensky), file=qif_file)
            print("T{:.2f}".format(order.total_payment), file=qif_file)
            # It appears that splits show up in GNUCash in opposite order.
            write_split(
                qif_file,
                order.opensky_commission,
                args.acct_commission
            )
            write_split(
                qif_file,
                order.cc_processing,
                args.acct_cc_processing
            )
            write_split(
                qif_file,
                order.restocking_fee,
                args.acct_restocking
            )
            write_split(
                qif_file,
                order.sales_tax,
                args.acct_sales_tax
            )
            write_split(
                qif_file,
                order.opensky_credits,
                args.acct_credits
            )
            write_split(
                qif_file,
                order.shipping,
                args.acct_shipping
            )
            write_split(
                qif_file,
                order.item_price,
                args.acct_sales,
                order.skus
            )
            print("^", file=qif_file)

def main():
    """main"""

    arg_parser = argparse.ArgumentParser(
        description="Converts an OpenSky .csv file to a Quicken .qif file."
    )
    arg_parser.add_argument(
        "--acct-cc_processing",
        metavar="NAME",
        default="Expenses:CC Processing Fees",
        help="credit card processing fees expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-credits",
        metavar="NAME",
        default="Expenses:OpenSky Credits",
        help="OpenSky credits expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-commission",
        metavar="NAME",
        default="Expenses:OpenSky Commissions",
        help="OpenSky commissions expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-opensky",
        metavar="NAME",
        default="Assets:OpenSky",
        help="OpenSky asset account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-shipping",
        metavar="NAME",
        default="Expenses:Postage and Delivery",
        help="shipping expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-restocking",
        metavar="NAME",
        default="Expenses:Restocking Fees",
        help="restocking fees expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-sales",
        metavar="NAME",
        default="Income:Sales - OpenSky",
        help="sales income account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--acct-sales-tax",
        metavar="NAME",
        default="Expenses:Sales Tax",
        help="sales tax expense account (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )
    arg_parser.add_argument(
        "csvfile",
        help="input CSV filename"
    )
    arg_parser.add_argument(
        "qiffile",
        nargs="?",
        help="output QIF filename"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()
    if args.qiffile:
        if args.csvfile == args.qiffile:
            arg_parser.error("csvfile and qiffile are the same file")
    else:
        (root, ext) = os.path.splitext(args.csvfile)
        args.qiffile = root + ".qif"

    # Read a .csv file exported from OpenSky.
    orders = read_opensky(args)

    # Check that the numbers add up for each order.
    validate_orders(orders)

    # Write a .qif file.
    write_qif(args, orders)

    return 0

if __name__ == "__main__":
    sys.exit(main())
