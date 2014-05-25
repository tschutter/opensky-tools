#!/usr/bin/env python

"""
Each row in a payments CSV file exported from OpenSky corresponds to a
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
import Tkinter
import tkFileDialog
import tkMessageBox


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

    def validate(self, ui, order_id):
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
            ui.warning(
                (
                    "For Order ID = {},"
                    " Item price {:.2f} != calculated price {:.2f}"
                ).format(
                    order_id,
                    self.item_price,
                    item_price
                )
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


def read_opensky(ui):
    """Read a .csv file exported from OpenSky."""

    orders = {}
    try:
        with open(ui.args.csvfile) as opensky_file:
            opensky_reader = csv.DictReader(opensky_file)
            for rownum, fields in enumerate(opensky_reader):
                try:
                    order_id = fields["Order ID"]
                    sku = fields["SKU"]
                    item_price = get_float(fields, "Item price")
                    shipping = get_float(fields, "Shipping price")
                    # Sign of credits is inverted in exported .csv file.
                    opensky_credits = -get_float(fields, "Credits")
                    sales_tax = get_float(fields, "Sales tax")
                    restocking_fee = get_float(fields, "Restocking fee")
                    cc_processing = get_float(fields, "Credit card processing")
                    opensky_commission = get_float(
                        fields,
                        "OpenSky commission"
                    )
                    total_payment = get_float(fields, "Total payment")
                    date = fields["Original order date"]
                except KeyError as kex:
                    ui.fatal(
                        (
                            "Unable to find {} column in {}."
                            " Is this an OpenSky payments file?"
                        ).format(kex, ui.args.csvfile)
                    )
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
    except IOError:
        ui.fatal("Unable to open {} for reading".format(ui.args.csvfile))

    return orders


def validate_orders(ui, orders):
    """Check that the numbers add up for each order."""
    for order_id in orders:
        orders[order_id].validate(ui, order_id)


def write_split(qif_file, amount, account, memo=None):
    """Write a single split to a .qif file."""
    if amount != 0.0:
        print("S{}".format(account), file=qif_file)
        if memo:
            print("E{}".format(memo), file=qif_file)
        print("${:.2f}".format(amount), file=qif_file)


def write_qif(ui, orders):
    """Write a .qif file."""
    try:
        with open(ui.args.qiffile, "w") as qif_file:
            print("!Account", file=qif_file)
            print("N{}".format(ui.args.acct_opensky), file=qif_file)
            print("TBank", file=qif_file)
            print("^", file=qif_file)
            print("!Type:Bank", file=qif_file)
            for order_id in sorted(orders):
                order = orders[order_id]
                print("D{}".format(order.date), file=qif_file)
                print("POpenSky order {}".format(order_id), file=qif_file)
                print("L{}".format(ui.args.acct_opensky), file=qif_file)
                print("T{:.2f}".format(order.total_payment), file=qif_file)
                # It appears that splits show up in GNUCash in opposite order.
                write_split(
                    qif_file,
                    order.opensky_commission,
                    ui.args.acct_commission
                )
                write_split(
                    qif_file,
                    order.cc_processing,
                    ui.args.acct_cc_processing
                )
                write_split(
                    qif_file,
                    order.restocking_fee,
                    ui.args.acct_restocking
                )
                write_split(
                    qif_file,
                    order.sales_tax,
                    ui.args.acct_sales_tax
                )
                write_split(
                    qif_file,
                    order.opensky_credits,
                    ui.args.acct_credits
                )
                write_split(
                    qif_file,
                    order.shipping,
                    ui.args.acct_shipping
                )
                write_split(
                    qif_file,
                    order.item_price,
                    ui.args.acct_sales,
                    order.skus
                )
                print("^", file=qif_file)
    except IOError:
        ui.fatal("Unable to open {} for writing".format(ui.args.qiffile))


class AppUI(object):
    def __init__(self, args):
        self.args = args

    def convert(self):
        """Perform the actual conversion."""

        # Read a .csv file exported from OpenSky.
        orders = read_opensky(self)

        # Check that the numbers add up for each order.
        validate_orders(self, orders)

        # Write a .qif file.
        write_qif(self, orders)


class AppGUI(AppUI):
    def __init__(self, args, root):
        AppUI.__init__(self, args)
        self.root = root
        self.frame = Tkinter.Frame(root)
        self.frame.pack()

        # CSV file
        csv_group = Tkinter.LabelFrame(
            self.frame,
            text="CSV file",
            padx=5,
            pady=5
        )
        csv_group.pack(padx=10, pady=10)
        self.csv_entry = Tkinter.Entry(csv_group, width=60)
        self.csv_entry.pack(side=Tkinter.LEFT)
        if args.csvfile:
            self.csv_entry.insert(0, args.csvfile)
        select_csvfile_button = Tkinter.Button(
            csv_group,
            text="Select",
            command=self.select_csvfile
        )
        select_csvfile_button.pack(side=Tkinter.LEFT)

        # QIF file
        qif_group = Tkinter.LabelFrame(
            self.frame,
            text="QIF file",
            padx=5,
            pady=5
        )
        qif_group.pack(padx=10, pady=10)
        self.qif_entry = Tkinter.Entry(qif_group, width=60)
        self.qif_entry.pack(side=Tkinter.LEFT)
        if args.qiffile:
            self.qif_entry.insert(0, args.qiffile)
        select_qiffile_button = Tkinter.Button(
            qif_group,
            text="Select",
            command=self.select_qiffile
        )
        select_qiffile_button.pack(side=Tkinter.LEFT)

        # Convert button
        self.convert_button = Tkinter.Button(
            self.frame,
            text="Convert",
            command=self.convert_action,
            default=Tkinter.ACTIVE
        )
        self.convert_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Cancel button
        self.cancel_button = Tkinter.Button(
            self.frame,
            text="Cancel",
            command=self.frame.quit
        )
        self.cancel_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

    def set_qiffile(self, qiffile):
        self.args.qiffile = qiffile
        self.qif_entry.delete(0, Tkinter.END)
        self.qif_entry.insert(0, qiffile)

    def set_csvfile(self, csvfile):
        self.args.csvfile = csvfile
        self.csv_entry.delete(0, Tkinter.END)
        self.csv_entry.insert(0, csvfile)
        root, ext = os.path.splitext(csvfile)
        self.set_qiffile(root + ".qif")

    def select_csvfile(self):
        self.args.csvfile = self.csv_entry.get()
        csvfile = tkFileDialog.askopenfilename(
            title="Choose .csv file",
            defaultextension=".csv",
            filetypes=[("Comma Separated Value", "*.csv")],
            initialfile=self.args.csvfile,
            multiple=False
        )
        if csvfile:
            self.set_csvfile(csvfile)

    def select_qiffile(self):
        self.args.qiffile = self.qif_entry.get()
        qiffile = tkFileDialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".qif",
            filetypes=[("Quicken Interchange Format", "*.qif")],
            title="Output .qif file",
            initialfile=self.args.qiffile
        )
        if qiffile:
            self.set_qiffile(qiffile)

    def error(self, msg):
        tkMessageBox.showerror("Error", msg)

    def fatal(self, msg):
        self.error(msg)
        self.frame.quit()
        sys.exit(1)

    def warning(self, msg):
        tkMessageBox.showwarning("Warning", msg)

    def convert_action(self):
        self.args.csvfile = self.csv_entry.get()
        self.args.qiffile = self.qif_entry.get()
        if not self.args.csvfile:
            tkMessageBox.showerror("Error", "CSV file not specified")
        elif not self.args.qiffile:
            tkMessageBox.showerror("Error", "QIF file not specified")
        else:
            self.convert()
            self.frame.quit()


class AppCLI(AppUI):
    def __init__(self, args):
        AppUI.__init__(self, args)

    def error(self, msg):
        print("ERROR: {}".format(msg), file=sys.stderr)

    def fatal(self, msg):
        self.error(msg)
        sys.exit(1)

    def warning(self, msg):
        print("WARNING: {}".format(msg), file=sys.stderr)


def main():
    """main"""

    arg_parser = argparse.ArgumentParser(
        description="Converts an OpenSky .csv file to a Quicken .qif file."
    )
    arg_parser.add_argument(
        "--gui",
        action="store_true",
        default=False,
        help="display a GUI"
    )
    arg_parser.add_argument(
        "--acct-cc_processing",
        metavar="NAME",
        default="Expenses:CC Processing Fees",
        help=(
            "credit card processing fees expense account"
            " (default=%(default)s)"
        )
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
        "csvfile",
        nargs="?",
        help="input CSV filename"
    )
    arg_parser.add_argument(
        "qiffile",
        nargs="?",
        help="output QIF filename"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()
    if args.csvfile:
        if args.qiffile:
            if args.csvfile == args.qiffile:
                arg_parser.error("csvfile and qiffile are the same file")
        else:
            (root, ext) = os.path.splitext(args.csvfile)
            args.qiffile = root + ".qif"
    else:
        if not args.gui:
            arg_parser.error("csvfile not specified")

    if args.gui:
        root = Tkinter.Tk()
        root.title("opensky2qif")
        AppGUI(args, root)
        root.mainloop()
        root.destroy()
    else:
        cli_ui = AppCLI(args)
        cli_ui.convert()

    return 0


if __name__ == "__main__":
    sys.exit(main())
