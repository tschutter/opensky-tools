#!/usr/bin/env python

"""Converts a payment-details CSV (Comma Separated Value) file
exported from the OpenSky Dashboard -> Accounts -> Payments to a QIF
(Quicken Interchange Format) file.

Each row in the CSV file corresponds to a line item in an order.
Lines with the same "Order ID" are combined into a single order.
Lines with the same "Payment date" are combined into a single payment.

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


class Order(object):
    """A single user order."""
    def __init__(
        self,
        order_date,
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
        self.order_date = order_date
        self.skus = [sku]
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
        """Update the values with another line item row with the same
        'Order Id'.
        """
        if sku not in self.skus:
            self.skus.append(sku)
        self.item_price += item_price
        self.shipping += shipping
        self.opensky_credits += opensky_credits
        self.sales_tax += sales_tax
        self.restocking_fee += restocking_fee
        self.cc_processing += cc_processing
        self.opensky_commission += opensky_commission
        self.total_payment += total_payment

    def validate(self, app_ui, order_id):
        """Emit a warning if the values to not add up."""
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
        if round(self.item_price, 2) != item_price:
            app_ui.warning(
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


class Payment(object):
    """A single payment from OpenSky."""
    def __init__(self, order_id, total_payment):
        self.order_ids = [order_id]
        self.total_payment = total_payment

    def update(self, order_id, total_payment):
        """Update the values with another line item row with the same
        'Payment date'.
        """
        if order_id not in self.order_ids:
            self.order_ids.append(order_id)
        self.total_payment += total_payment


def get_float(fields, key):
    """Convert a string value to a float, handling blank values."""
    value = fields[key]
    try:
        value = float(value)
    except ValueError:
        value = 0.0
    return value


def normalize_date(date):
    """
    Return date in YYYY-MM-DD form.

    This prevents GnuCash from asking you to "set a date format for
    this QIF file" if the dates in the file are ambiguous.

    Prior to 2015-07-25 the payment date was YYYY-MM-DD and the order
    date was MM/DD/YYYY.  Therefore we can make no assumptions about
    the format of dates in the .csv file.
    """

    if "/" in date:
        # Assume MM/DD/YYYY.
        date = "{0[2]}-{0[0]}-{0[1]}".format(date.split("/"))

    return date


def read_opensky(app_ui):
    """Read a .csv file exported from OpenSky."""

    orders = {}
    payments = {}
    try:
        with open(app_ui.args.csvfile) as opensky_file:
            opensky_reader = csv.DictReader(opensky_file)
            for fields in opensky_reader:
                # Get information from line item row.
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
                    order_date = normalize_date(fields["Original order date"])
                    payment_date = normalize_date(fields["Payment date"])
                except KeyError as kex:
                    app_ui.fatal(
                        (
                            "Unable to find {} column in {}."
                            " Is this an OpenSky payments file?"
                        ).format(kex, app_ui.args.csvfile)
                    )

                # Update orders.
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
                        order_date,
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

                # Update payments.
                if payment_date in payments:
                    payments[payment_date].update(order_id, total_payment)
                else:
                    payments[payment_date] = Payment(order_id, total_payment)

    except IOError:
        app_ui.fatal(
            "Unable to open {} for reading".format(app_ui.args.csvfile)
        )

    if len(orders) == 0:
        app_ui.fatal(
            "No data found in {}".format(app_ui.args.csvfile)
        )

    return (orders, payments)


def validate_orders(app_ui, orders):
    """Check that the numbers add up for each order."""
    for order_id in orders:
        orders[order_id].validate(app_ui, order_id)


def write_split(qif_file, amount, account, memo=None):
    """Write a single split to a .qif file."""
    if amount != 0.0:
        print("S{}".format(account), file=qif_file)
        if memo:
            print("E{}".format(memo), file=qif_file)
        print("${:.2f}".format(amount), file=qif_file)


def write_qif_header(app_ui, qif_file):
    """Write header to a .qif file."""
    print("!Account", file=qif_file)
    print("N{}".format(app_ui.args.acct_opensky), file=qif_file)
    print("TBank", file=qif_file)
    print("^", file=qif_file)
    print("!Type:Bank", file=qif_file)


def list_to_string(item_name, values):
    """Return a pretty version of the list."""
    values = sorted(values)
    if len(values) < 5:
        values_string = "{}{}={}".format(
            item_name,
            "s" if len(values) > 1 else "",
            ",".join(values)
        )
    else:
        values_string = "{} {}s from {} to {}".format(
            len(values),
            item_name,
            values[0],
            values[-1]
        )
    return values_string


def write_qif_orders(app_ui, orders, qif_file):
    """Write orders to a .qif file."""
    for order_id in sorted(orders):
        order = orders[order_id]
        print("D{}".format(order.order_date), file=qif_file)
        print("POpenSky order {}".format(order_id), file=qif_file)
        print("L{}".format(app_ui.args.acct_opensky), file=qif_file)
        print("T{:.2f}".format(order.total_payment), file=qif_file)
        # It appears that splits show up in GNUCash in the
        # opposite order than they are specified in the
        # imported file.
        write_split(
            qif_file,
            order.opensky_commission,
            app_ui.args.acct_commission
        )
        write_split(
            qif_file,
            order.cc_processing,
            app_ui.args.acct_cc_processing
        )
        write_split(
            qif_file,
            order.restocking_fee,
            app_ui.args.acct_restocking
        )
        write_split(
            qif_file,
            order.sales_tax,
            app_ui.args.acct_sales_tax
        )
        write_split(
            qif_file,
            order.opensky_credits,
            app_ui.args.acct_credits
        )
        write_split(
            qif_file,
            order.shipping,
            app_ui.args.acct_shipping
        )
        write_split(
            qif_file,
            order.item_price,
            app_ui.args.acct_sales,
            list_to_string("SKU", order.skus)
        )
        print("^", file=qif_file)


def write_qif_payments(app_ui, payments, qif_file):
    """Write payments to a .qif file."""
    for payment_date in sorted(payments):
        payment = payments[payment_date]
        print("D{}".format(payment_date), file=qif_file)
        print("POpenSky payment", file=qif_file)
        print("L{}".format(app_ui.args.acct_opensky), file=qif_file)
        print("T{:.2f}".format(-payment.total_payment), file=qif_file)
        write_split(
            qif_file,
            payment.total_payment,
            app_ui.args.acct_deposit,
            "{}".format(list_to_string("Order ID", payment.order_ids)),
        )
        print("^", file=qif_file)


def write_qif(app_ui, orders, payments):
    """Write a .qif file."""
    try:
        with open(app_ui.args.qiffile, "w") as qif_file:
            write_qif_header(app_ui, qif_file)
            write_qif_orders(app_ui, orders, qif_file)
            write_qif_payments(app_ui, payments, qif_file)
    except IOError:
        app_ui.fatal(
            "Unable to open {} for writing".format(app_ui.args.qiffile)
        )


class AppUI(object):
    """Base application user interface."""
    def __init__(self, args):
        self.args = args

    def convert(self):
        """Perform the actual conversion."""

        # Read a .csv file exported from OpenSky.
        orders, payments = read_opensky(self)

        # Check that the numbers add up for each order.
        validate_orders(self, orders)

        # Write a .qif file.
        write_qif(self, orders, payments)


class AppGUI(AppUI):
    """Application graphical user interface."""

    # Width of filename text entry boxes.
    FILENAME_WIDTH = 95

    # pylint: disable=no-self-use
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
        self.csv_entry = Tkinter.Entry(csv_group, width=AppGUI.FILENAME_WIDTH)
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
        self.qif_entry = Tkinter.Entry(qif_group, width=AppGUI.FILENAME_WIDTH)
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

        # Quit button
        self.cancel_button = Tkinter.Button(
            self.frame,
            text="Quit",
            command=self.frame.quit
        )
        self.cancel_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

    def set_csvfile(self, csvfile):
        """Set the name of the CSV file."""
        self.args.csvfile = csvfile
        self.csv_entry.delete(0, Tkinter.END)
        self.csv_entry.insert(0, csvfile)
        if len(csvfile) > 0:
            root = os.path.splitext(csvfile)[0]
            qiffile = root + ".qif"
        else:
            qiffile = ""
        self.set_qiffile(qiffile)

    def set_qiffile(self, qiffile):
        """Set the name of the QIF file."""
        self.args.qiffile = qiffile
        self.qif_entry.delete(0, Tkinter.END)
        self.qif_entry.insert(0, qiffile)

    def select_csvfile(self):
        """Display a dialog to select a CSV file."""
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
        """Display a dialog to select a QIF file."""
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
        """Display an error message."""
        tkMessageBox.showerror("Error", msg)

    def fatal(self, msg):
        """Display an error message and exit."""
        self.error(msg)
        self.frame.quit()
        sys.exit(1)

    def warning(self, msg):
        """Display a warning message."""
        tkMessageBox.showwarning("Warning", msg)

    def convert_action(self):
        """Perform the conversion and exit."""
        self.args.csvfile = self.csv_entry.get()
        self.args.qiffile = self.qif_entry.get()
        if not self.args.csvfile:
            tkMessageBox.showerror("Error", "CSV file not specified")
        elif not self.args.qiffile:
            tkMessageBox.showerror("Error", "QIF file not specified")
        else:
            self.convert()
            self.set_csvfile("")

class AppCLI(AppUI):
    """Application command line user interface."""
    # pylint: disable=no-self-use
    def __init__(self, args):
        AppUI.__init__(self, args)

    def error(self, msg):
        """Display an error message."""
        print("ERROR: {}".format(msg), file=sys.stderr)

    def fatal(self, msg):
        """Display an error message and exit."""
        self.error(msg)
        sys.exit(1)

    def warning(self, msg):
        """Display a warning message."""
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
        "--acct-deposit",
        metavar="NAME",
        default="Assets:Checking",
        help="payment deposit account (default=%(default)s)"
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
            root = os.path.splitext(args.csvfile)[0]
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
