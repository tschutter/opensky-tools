opensky-tools
=============

opensky-tools is a set of utilities that work with OpenSky transaction
data.

opensky2qif.py
    Converts a payment-details CSV (Comma Separated Value) file
    exported from the OpenSky Dashboard -> Accounts -> Payments to a
    QIF (Quicken Interchange Format) file.  The default is a command
    line interface, but you can specify --gui to use a GUI interface.

    The output .csv file has been successfully imported into
    gnucash-2.6.1.  It has not been tested with any Quicken products.
