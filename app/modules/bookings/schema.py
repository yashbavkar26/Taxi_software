"""Schema and column mappings for the bookings table."""

from app.common.normalizer import ColumnDef
from app.common.schema_base import TableSchema

BOOKINGS_COLUMNS = {
    "received": ColumnDef(db_column="received", data_type="boolean", is_required=False),
    "sl_no": ColumnDef(db_column="sl_no", data_type="integer", is_required=False),
    "coordinator": ColumnDef(db_column="coordinator", data_type="string", is_required=False, max_length=255),
    "guest_name": ColumnDef(db_column="guest_name", data_type="string", is_required=False, max_length=255),
    "guest_mobile_no": ColumnDef(db_column="guest_mobile_no", data_type="phone", is_required=False, max_length=20),
    "travel_date": ColumnDef(db_column="travel_date", data_type="date", is_required=False),
    "pickup_location": ColumnDef(db_column="pickup_location", data_type="string", is_required=False, max_length=255),
    "drop_location": ColumnDef(db_column="drop_location", data_type="string", is_required=False, max_length=255),
    "start_time": ColumnDef(db_column="start_time", data_type="time", is_required=False),
    "return_time": ColumnDef(db_column="return_time", data_type="time", is_required=False),
    "reason": ColumnDef(db_column="reason", data_type="string", is_required=False),
    "driver_name": ColumnDef(db_column="driver_name", data_type="string", is_required=False, max_length=255),
    "contact": ColumnDef(db_column="contact", data_type="phone", is_required=False, max_length=20),
    "vehicle_type": ColumnDef(db_column="vehicle_type", data_type="string", is_required=False, max_length=255),
    "company": ColumnDef(db_column="company", data_type="string", is_required=False, max_length=255, default_value="GIM"),
}

BOOKINGS_HEADER_ALIASES = {
    "received": "received",
    "recived": "received",
    "sl.no.": "sl_no",
    "sl. no.": "sl_no",
    "sl.no": "sl_no",
    "sl no": "sl_no",
    "sl_no": "sl_no",
    "serial no": "sl_no",
    "sr. no": "sl_no",
    "sr no": "sl_no",
    "sr. no.": "sl_no",
    "2": "sl_no",
    "name": "coordinator",
    "coordinator": "coordinator",
    "coordinator name": "coordinator",
    "co odrinator": "coordinator",
    "co-ordinator": "coordinator",
    "coord": "coordinator",
    "group no": "coordinator",
    "group no.": "coordinator",
    "guest": "guest_name",
    "guest name": "guest_name",
    "group leader": "guest_name",
    "guest mobile no.": "guest_mobile_no",
    "guest mobile no": "guest_mobile_no",
    "guest mobile": "guest_mobile_no",
    "group leader mobile no.": "guest_mobile_no",
    "group leader mobile no": "guest_mobile_no",
    "group leader mobile": "guest_mobile_no",
    "travel date": "travel_date",
    "date": "travel_date",
    "from": "pickup_location",
    "pickup": "pickup_location",
    "pickup location": "pickup_location",
    "to": "drop_location",
    "drop": "drop_location",
    "drop location": "drop_location",
    "start time": "start_time",
    "return time": "return_time",
    "reason": "reason",
    "number of students": "reason",
    "no. of students": "reason",
    "number of cabs": "reason",
    "driver": "driver_name",
    "driver name": "driver_name",
    "contact": "contact",
    "driver contact": "contact",
    "vehicle type": "vehicle_type",
    "vehicle_type": "vehicle_type",
    "veh type": "vehicle_type",
    "vehicle": "vehicle_type",
    "vech type": "vehicle_type",
    "vech. type": "vehicle_type",
    "vech.type": "vehicle_type",
    "vech_type": "vehicle_type",
    "type": "vehicle_type",
    "vech no": "vehicle_type",
    "vech no.": "vehicle_type",
    "vech. no": "vehicle_type",
    "vech. no.": "vehicle_type",
    "vech.no": "vehicle_type",
    "vech_no": "vehicle_type",
    "vehicle no": "vehicle_type",
    "vehicle no.": "vehicle_type",
    "vehicle number": "vehicle_type",
    "vehicle_no": "vehicle_type",
    "company": "company",
    "company name": "company",
    "comp": "company",
    "vendor": "company",
    "vendor name": "company",
    "vend": "company",
}

CANONICAL_BOOKING_HEADERS = [
    "Received",
    "Sl.No.",
    "Name",
    "Guest",
    "Guest Mobile No.",
    "Travel Date",
    "From",
    "To",
    "Start Time",
    "Return Time",
    "Reason",
    "Driver",
    "Contact",
    "Vech No",
    "Company",
]


BOOKINGS_SCHEMA = TableSchema(
    table_name="bookings",
    primary_key="id",
    columns=BOOKINGS_COLUMNS,
    header_aliases=BOOKINGS_HEADER_ALIASES,
    required_headers={"travel_date", "driver_name"},
    generated_columns={
        "driver_message",
        "guest_message",
        "driver message",
        "guest message",
        "guest link",
        "driver link",
        "message",
        "guest text",
        "driver text",
    },
)


def format_guest_message(
    driver_name: str = "",
    contact: str = "",
    start_time_str: str = "",
    pickup_location: str = "",
    company: str = "GIM",
) -> str:
    """Format the guest notification message matching MySQL stored generated column."""
    comp = company or "GIM"
    driver = driver_name or ""
    phone = contact or ""
    time_val = start_time_str or ""
    pickup = pickup_location or ""
    return (
        f"From {comp}\n"
        f"Driver {driver} {phone} will come to pick you tomorrow at {time_val} at {pickup}"
    )


def format_driver_message(
    guest_name: str = "",
    guest_mobile_no: str = "",
    pickup_location: str = "",
    drop_location: str = "",
    start_time_str: str = "",
) -> str:
    """Format the driver notification message matching MySQL stored generated column."""
    g_name = guest_name or ""
    g_phone = guest_mobile_no or ""
    p_loc = pickup_location or ""
    d_loc = drop_location or ""
    time_val = start_time_str or ""
    return (
        f"Tomorrow\n"
        f"{g_name} {g_phone}\n"
        f"{p_loc} to {d_loc}\n"
        f"{time_val}"
    )

