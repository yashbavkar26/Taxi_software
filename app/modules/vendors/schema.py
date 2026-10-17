"""Schema and column mappings for the masterfile / vendor book table.

Matches the masterfile table schema in taxi_schema.sql and provides robust
header aliasing for Excel import/export pipelines.
"""

from typing import Dict, Set
from app.common.normalizer import ColumnDef
from app.common.schema_base import TableSchema

MASTERFILE_COLUMNS: Dict[str, ColumnDef] = {
    "sl_no": ColumnDef(db_column="sl_no", data_type="integer", is_required=False),
    "vendor": ColumnDef(db_column="vendor", data_type="string", is_required=True, max_length=255),
    "travel_date": ColumnDef(db_column="travel_date", data_type="date", is_required=False),
    "vehicle_no": ColumnDef(db_column="vehicle_no", data_type="string", is_required=False, max_length=50),
    "vehicle_type": ColumnDef(db_column="vehicle_type", data_type="string", is_required=False, max_length=100),
    "starting_km": ColumnDef(db_column="starting_km", data_type="string", is_required=False),
    "ending_km": ColumnDef(db_column="ending_km", data_type="string", is_required=False),
    "total_km": ColumnDef(db_column="total_km", data_type="string", is_required=False),
    "night": ColumnDef(db_column="night", data_type="string", is_required=False, default_value="0.00"),
    "rate": ColumnDef(db_column="rate", data_type="string", is_required=False, default_value="0.00"),
    "toll": ColumnDef(db_column="toll", data_type="string", is_required=False, default_value="0.00"),
    "amt": ColumnDef(db_column="amt", data_type="string", is_required=False, default_value="0.00"),
    "guest_name": ColumnDef(db_column="guest_name", data_type="string", is_required=False, max_length=255),
    "total_amt": ColumnDef(db_column="total_amt", data_type="string", is_required=False, default_value="0.00"),
    "payment_details": ColumnDef(db_column="payment_details", data_type="boolean", is_required=False, default_value=False),
    "in_time": ColumnDef(db_column="in_time", data_type="time", is_required=False),
    "out_time": ColumnDef(db_column="out_time", data_type="time", is_required=False),
    "driver_name": ColumnDef(db_column="driver_name", data_type="string", is_required=False, max_length=255),
    "places": ColumnDef(db_column="places", data_type="string", is_required=False, max_length=255),
    "payment_date": ColumnDef(db_column="payment_date", data_type="date", is_required=False),
}

MASTERFILE_HEADER_ALIASES: Dict[str, str] = {
    # sl_no
    "sl. no": "sl_no",
    "sl. no.": "sl_no",
    "sl.no": "sl_no",
    "sl.no.": "sl_no",
    "sl no": "sl_no",
    "sl_no": "sl_no",
    "serial no": "sl_no",
    "sr. no": "sl_no",
    "sr. no.": "sl_no",
    "sr no": "sl_no",
    "s.no": "sl_no",
    "s. no": "sl_no",
    # vendor
    "vendor": "vendor",
    "vendor name": "vendor",
    "vendor_name": "vendor",
    "agency": "vendor",
    # travel_date
    "date": "travel_date",
    "date ": "travel_date",
    "travel date": "travel_date",
    "travel_date": "travel_date",
    "trip date": "travel_date",
    # vehicle_no
    "vech.no": "vehicle_no",
    "vech. no": "vehicle_no",
    "vech no": "vehicle_no",
    "vech_no": "vehicle_no",
    "vehicle no": "vehicle_no",
    "vehicle no.": "vehicle_no",
    "vehicle_no": "vehicle_no",
    "car no": "vehicle_no",
    "reg no": "vehicle_no",
    # vehicle_type
    "vech type": "vehicle_type",
    "vech. type": "vehicle_type",
    "vehicle type": "vehicle_type",
    "vehicle_type": "vehicle_type",
    "car type": "vehicle_type",
    "model": "vehicle_type",
    # starting_km
    "staring km": "starting_km",  # common typo in Masterfile.xlsx
    "staring_km": "starting_km",
    "starting km": "starting_km",
    "starting_km": "starting_km",
    "start km": "starting_km",
    "starting reading": "starting_km",
    "start reading": "starting_km",
    "start_km": "starting_km",
    # ending_km
    "ending km": "ending_km",
    "ending_km": "ending_km",
    "end km": "ending_km",
    "ending reading": "ending_km",
    "end reading": "ending_km",
    "end_km": "ending_km",
    # total_km
    "total km": "total_km",
    "total kms": "total_km",
    "total_km": "total_km",
    "total_kms": "total_km",
    "kms": "total_km",
    # night
    "night": "night",
    "night charges": "night",
    "night_charges": "night",
    "night charge": "night",
    # rate
    "rate": "rate",
    "rate per km": "rate",
    "rate per km above 80": "rate",
    "rate_per_km_above_80": "rate",
    # toll
    "toll": "toll",
    "toll charges": "toll",
    "toll_charges": "toll",
    # amt
    "amt": "amt",
    "amt.": "amt",
    "amount": "amt",
    "total km rs": "amt",
    "total_km_rs": "amt",
    "base amt": "amt",
    # guest_name
    "guest": "guest_name",
    "guest ": "guest_name",
    "guest name": "guest_name",
    "guest_name": "guest_name",
    "passenger": "guest_name",
    # total_amt
    "total amt": "total_amt",
    "total amount": "total_amt",
    "total rs": "total_amt",
    "total_amt": "total_amt",
    "gross amt": "total_amt",
    "net amt": "total_amt",
    # payment_details
    "payment details": "payment_details",
    "payment_details": "payment_details",
    "payment status": "payment_details",
    "payment": "payment_details",
    "paid": "payment_details",
    # in_time
    "in time": "in_time",
    "in_time": "in_time",
    "start time": "in_time",
    "pickup time": "in_time",
    # out_time
    "out time": "out_time",
    "out_time": "out_time",
    "drop time": "out_time",
    "return time": "out_time",
    # driver_name
    "driver name": "driver_name",
    "driver_name": "driver_name",
    "driver": "driver_name",
    # places
    "places": "places",
    "place": "places",
    "from - to": "places",
    "from -  to": "places",
    "from_to": "places",
    "route": "places",
    "destination": "places",
    # payment_date
    "payment date": "payment_date",
    "payment_date": "payment_date",
    "paid date": "payment_date",
    "date of payment": "payment_date",
}

MASTERFILE_REQUIRED_HEADERS: Set[str] = {"vendor"}

MASTERFILE_GENERATED_COLUMNS: Set[str] = {"total_km", "total_amt"}

MASTERFILE_SCHEMA = TableSchema(
    table_name="masterfile",
    primary_key="id",
    columns=MASTERFILE_COLUMNS,
    header_aliases=MASTERFILE_HEADER_ALIASES,
    required_headers=MASTERFILE_REQUIRED_HEADERS,
    generated_columns=MASTERFILE_GENERATED_COLUMNS,
)


def is_masterfile_sheet(sheet_name: str) -> bool:
    """Return True if the sheet name indicates masterfile / vendor records."""
    cleaned = sheet_name.strip().lower()
    return cleaned in (
        "masterfile",
        "master file",
        "master",
        "vendor book",
        "vendorbook",
        "vendor",
        "vendors",
        "sheet1",
    )
