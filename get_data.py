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



# New function: connection_type
def connection_type(csv_filename="connections.csv", subfolder="program files"):
    """
    Reads the connections CSV file and returns a sorted list of unique connection types.

    Args:
        csv_filename (str): CSV filename in the `program files` subdirectory.
        subfolder (str): Subdirectory under this file's directory.

    Returns:
        list: Sorted list of unique connection types found in the CSV.

    Raises:
        FileNotFoundError: If the CSV file does not exist.
        ValueError: If the CSV file does not contain a 'type' column.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)

    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Connections file not found: {csv_path}")

    types = set()
    with open(csv_path, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or '\ufeffنوع' not in reader.fieldnames:
            raise ValueError(f"'نوع' column not found in CSV: {csv_path}")
        for row in reader:
            value = row.get('\ufeffنوع')
            if value:
                types.add(value.strip())
    return sorted(types)


# New function: products_for_connection_type
def products_for_connection_type(type_value, csv_filename="connections.csv", subfolder="program files"):
    """
    Returns a sorted list of unique products ('محصول') for a given connection type ('نوع').

    Args:
        type_value (str): The value of the connection type to filter on.
        csv_filename (str): The CSV file name.
        subfolder (str): Subdirectory for the CSV.

    Returns:
        list: Sorted list of unique products under the given type.

    Raises:
        FileNotFoundError: If the CSV does not exist.
        ValueError: If the required columns are not found.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Connections file not found: {csv_path}")

    products = set()
    with open(csv_path, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        # Must use BOM for first field
        if not reader.fieldnames or '\ufeffنوع' not in reader.fieldnames or 'محصول' not in reader.fieldnames:
            raise ValueError(f"'نوع' or 'محصول' column not found in CSV: {csv_path}")
        for row in reader:
            if row.get('\ufeffنوع', '').strip() == type_value.strip():
                prod = row.get('محصول')
                if prod:
                    products.add(prod.strip())
    return sorted(products)


# New function: sizes_for_type_and_product
def sizes_for_type_and_product(type_value, product_value, csv_filename="connections.csv", subfolder="program files"):
    """
    Returns a sorted list of unique sizes ('اندازه (mm)') for a given type and product.

    Args:
        type_value (str): The value for the '\ufeffنوع' column.
        product_value (str): The value for the 'محصول' column.
        csv_filename (str): The CSV file name.
        subfolder (str): The folder for the CSV file.

    Returns:
        list: Sorted list of unique sizes for the type and product.

    Raises:
        FileNotFoundError: If the CSV is missing.
        ValueError: If columns are not found.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Connections file not found: {csv_path}")

    sizes = set()
    with open(csv_path, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if (
            not reader.fieldnames or
            '\ufeffنوع' not in reader.fieldnames or
            'محصول' not in reader.fieldnames or
            'اندازه (mm)' not in reader.fieldnames
        ):
            raise ValueError(f"Expected columns not found in CSV: {csv_path}")
        for row in reader:
            if (
                row.get('\ufeffنوع', '').strip() == type_value.strip() and
                row.get('محصول', '').strip() == product_value.strip()
            ):
                size = row.get('اندازه (mm)')
                if size:
                    sizes.add(size.strip())
    # Optionally, convert to float for numeric sort, else return as string sort
    try:
        return sorted(sizes, key=lambda x: float(x.replace(',', '').replace(' ', '')))
    except Exception:
        return sorted(sizes)


# New function: row_for_type_product_size
def row_for_type_product_size(type_value, product_value, size_value, csv_filename="connections.csv", subfolder="program files"):
    """
    Returns a dict of all data for the row matching the given type, product, and size.

    Args:
        type_value (str): The value for the '\ufeffنوع' column.
        product_value (str): The value for the 'محصول' column.
        size_value (str): The value for the 'اندازه (mm)' column.
        csv_filename (str): The CSV file name.
        subfolder (str): The folder for the CSV file.

    Returns:
        dict: The row data as a dict if found, else None.

    Raises:
        FileNotFoundError: If the CSV is missing.
        ValueError: If columns are not found.
    """
    base_dir = os.path.dirname(__file__)
    csv_path = os.path.join(base_dir, subfolder, csv_filename)
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Connections file not found: {csv_path}")

    with open(csv_path, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if (
            not reader.fieldnames or
            '\ufeffنوع' not in reader.fieldnames or
            'محصول' not in reader.fieldnames or
            'اندازه (mm)' not in reader.fieldnames
        ):
            raise ValueError(f"Expected columns not found in CSV: {csv_path}")
        for row in reader:
            if (
                row.get('\ufeffنوع', '').strip() == type_value.strip() and
                row.get('محصول', '').strip() == product_value.strip() and
                row.get('اندازه (mm)', '').strip() == size_value.strip()
            ):
                # Return the row as a dict (with trimmed values)
                return {k: (v.strip() if isinstance(v, str) else v) for k, v in row.items()}
    return None

