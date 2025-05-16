import os
import csv

def get_sdr_for(pipe_grade, pn, csv_filename="pipe_series_sdr.csv", subfolder="program files"):
    """
    Read the pipe series SDR CSV and return the SDR corresponding to a given pipe grade and PN.

    Args:
        pipe_grade (str): Pipe grade (e.g., "PE63", "PE80", "PE100").
        pn (float or int): Nominal pressure rating (e.g., 2.0, 2.5, 3.2, etc.).
        csv_filename (str): CSV filename in the `program files` subdirectory.
        subfolder (str): Subdirectory under this file's directory, default "program files".

    Returns:
        float: SDR value for the specified pipe grade and PN.

    Raises:
        ValueError: If the CSV is empty, values are missing, or inputs are invalid.
        KeyError: If the pipe grade or PN is not found in the CSV.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)

    with open(csv_path, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        raise ValueError(f"{csv_path} is empty")

    # Parse SDR values from header row (skip the first empty column)
    header = rows[0]
    sdr_list = []
    for val in header[1:]:
        if val:
            try:
                sdr_list.append(float(val))
            except ValueError:
                raise ValueError(f"Invalid SDR value in header: {val!r}")

    try:
        pn = float(pn)
    except ValueError:
        raise ValueError(f"Invalid PN input: {pn!r}")

    # Locate the row for the requested pipe grade
    target_grade = str(pipe_grade).strip().upper()
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        if row[0].strip().upper() == target_grade:
            # Search each SDR column for matching PN
            for idx, cell in enumerate(row[1:]):
                if not cell:
                    continue
                try:
                    cell_pn = float(cell)
                except ValueError:
                    continue
                if cell_pn == pn:
                    return sdr_list[idx]
            # PN not found for this grade
            raise KeyError(f"PN {pn} not found for grade {pipe_grade} in CSV: {csv_path}")

    # Pipe grade not found
    raise KeyError(f"Pipe grade {pipe_grade} not found in CSV: {csv_path}")

def get_pn_for(pipe_grade, sdr, csv_filename="pipe_series_sdr.csv", subfolder="program files"):
    """
    Read the pipe series SDR CSV and return the PN corresponding to a given pipe grade and SDR.

    Args:
        pipe_grade (str): Pipe grade (e.g., "PE63", "PE80", "PE100").
        sdr (float or int): SDR value (e.g., 17, 11, etc.).
        csv_filename (str): CSV filename in the `program files` subdirectory.
        subfolder (str): Subdirectory under this file's directory, default "program files".

    Returns:
        float: PN value for the specified pipe grade and SDR.

    Raises:
        ValueError: If the CSV is empty, values are missing, or inputs are invalid.
        KeyError: If the pipe grade or SDR is not found in the CSV.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)

    with open(csv_path, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        raise ValueError(f"{csv_path} is empty")

    # Parse SDR list from header row (skip the first empty column)
    header = rows[0]
    sdr_list = []
    for val in header[1:]:
        if val:
            try:
                sdr_list.append(float(val))
            except ValueError:
                raise ValueError(f"Invalid SDR value in header: {val!r}")

    try:
        sdr = float(sdr)
    except ValueError:
        raise ValueError(f"Invalid SDR input: {sdr!r}")

    if sdr not in sdr_list:
        raise KeyError(f"SDR {sdr} not found in CSV: {csv_path}")

    col_idx = sdr_list.index(sdr) + 1

    target_grade = str(pipe_grade).strip().upper()
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        if row[0].strip().upper() == target_grade:
            if col_idx < len(row) and row[col_idx]:
                try:
                    return float(row[col_idx])
                except ValueError:
                    raise ValueError(f"Invalid PN value: {row[col_idx]!r}")
            else:
                raise ValueError(f"PN for grade {pipe_grade} and SDR {sdr} missing in CSV: {csv_path}")

    raise KeyError(f"Pipe grade {pipe_grade} not found in CSV: {csv_path}")

def load_weight_table(diameter, sdr, csv_filename="DIN_pivot.csv", subfolder="program files"):
    """
    Read the SDR vs. diameter weight table from a CSV and return the weight for a given diameter and SDR.

    Args:
        diameter (float): Diameter in mm.
        sdr (float): SDR value.
        csv_filename (str): CSV filename in the `program files` directory.
        subfolder (str): Subdirectory under this file's directory.

    Returns:
        float: Weight corresponding to the given diameter and SDR.

    Raises:
        ValueError: If the CSV is empty, or inputs are invalid, or the weight cell is missing or malformed.
        KeyError: If the diameter or SDR is not found in the CSV.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)

    with open(csv_path, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        raise ValueError(f"{csv_path} is empty")

    # Parse SDR list from header row (skip first label column)
    header = rows[0]
    sdr_list = []
    for val in header[1:]:
        if val:
            try:
                sdr_list.append(float(val))
            except ValueError:
                raise ValueError(f"Invalid SDR in header: {val!r}")

    # Validate inputs
    try:
        diameter = float(diameter)
    except ValueError:
        raise ValueError(f"Invalid diameter input: {diameter!r}")
    try:
        sdr = float(sdr)
    except ValueError:
        raise ValueError(f"Invalid SDR input: {sdr!r}")

    if sdr not in sdr_list:
        raise KeyError(f"SDR {sdr} not found in CSV: {csv_path}")

    col_idx = sdr_list.index(sdr) + 1

    # Locate row for the given diameter
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        try:
            row_d = float(row[0])
        except ValueError:
            continue
        if row_d == diameter:
            if col_idx < len(row) and row[col_idx]:
                try:
                    return float(row[col_idx])
                except ValueError:
                    raise ValueError(f"Invalid weight value: {row[col_idx]!r}")
            else:
                raise ValueError(f"Weight for diameter {diameter} and SDR {sdr} missing in CSV: {csv_path}")

    raise KeyError(f"Diameter {diameter} not found in CSV: {csv_path}")



def get_discount(order_price, csv_filename="discount.csv", subfolder="program files"):
    """
    Read the discount thresholds from a CSV and return the discount percentage
    applicable to the given order_price.

    CSV format in 'program files/discount.csv':
        price, percentage
    Example rows:
        100,2
        500,4
        700,6

    Args:
        order_price (float): Total price of the order.
        csv_filename (str): CSV filename in the `program files` subdirectory.
        subfolder (str): Subdirectory under this file's directory, default "program files".

    Returns:
        float: The discount percentage for the order_price (0.0 if no threshold met).

    Raises:
        FileNotFoundError: If the CSV file does not exist.
        ValueError: If the CSV is empty or contains invalid numeric values.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)

    # Read and parse thresholds
    thresholds = []
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Discount file not found: {csv_path}")
    with open(csv_path, newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 2 or not row[0] or not row[1]:
                continue
            try:
                price = float(row[0])
                pct = float(row[1])
            except ValueError:
                raise ValueError(f"Invalid discount entry: {row}")
            thresholds.append((price, pct))

    if not thresholds:
        raise ValueError(f"No discount thresholds found in: {csv_path}")

    # Sort thresholds by price ascending
    thresholds.sort(key=lambda x: x[0])

    # Determine applicable discount
    discount_pct = 0.0
    for price, pct in thresholds:
        if order_price >= price:
            discount_pct = pct
        else:
            break

    return discount_pct

