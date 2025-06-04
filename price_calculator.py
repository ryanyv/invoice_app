import os
from get_data import get_pn_for, get_sdr_for, load_weight_table, get_discount


def calculate_total_mass(length_meters, diameter, sdr):
    """
    Calculate the total pipe mass in kilograms given pipe length (meters), diameter, and SDR.

    Args:
        length_meters (float or int): Length of the pipe in meters.
        diameter (float or int): Pipe diameter in mm.
        sdr (float or int): SDR value.

    Returns:
        float: Total mass in kilograms.

    Raises:
        ValueError: If length_meters cannot be converted to float.
        KeyError/ValueError: Propagated from load_weight_table for invalid diameter or SDR.
    """
    # Get mass per meter (kg/m) from weight table
    mass_per_meter = load_weight_table(diameter, sdr)

    # Validate and convert length
    try:
        length = float(length_meters)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid pipe length: {length_meters!r}")

    # Compute and return total mass
    return mass_per_meter * length


def calculate_price(length_meters, diameter, sdr, price_per_kg):
    """
    Calculate the total price given pipe length, diameter, SDR, and price per kilogram.

    Args:
        length_meters (float or int): Pipe length in meters.
        diameter (float or int): Pipe diameter in mm.
        sdr (float or int): SDR value.
        price_per_kg (float or int): Price per kilogram.

    Returns:
        float: Total price.

    Raises:
        ValueError: If price_per_kg cannot be converted to float, or if length_meters is invalid (propagated).
        KeyError/ValueError: Propagated from calculate_total_mass for invalid diameter or SDR.
    """
    # Calculate mass first (kg)
    mass = calculate_total_mass(length_meters, diameter, sdr)

    # Validate and convert price
    try:
        price_per_kg = float(price_per_kg)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid price per kilogram: {price_per_kg!r}")

    # Compute and return total price
    return mass * price_per_kg


def calculate_price_per_kg_from_total(total_price, length_meters, diameter, sdr):
    """
    Calculate the price per kilogram from a total price and pipe specification.

    This is the inverse operation of ``calculate_price`` and therefore requires
    the pipe length in meters to derive the total mass used in the calculation.

    Args:
        total_price (float or int): Total price for ``length_meters`` of pipe.
        length_meters (float or int): Length of the pipe in meters.
        diameter (float or int): Pipe diameter in mm.
        sdr (float or int): SDR value.

    Returns:
        float: Price per kilogram.

    Raises:
        ValueError: If inputs cannot be converted to float, or if the calculated
            mass is zero/invalid.
        KeyError/ValueError: Propagated from ``calculate_total_mass`` for
            invalid diameter or SDR.
    """

    # Validate and convert numeric inputs
    try:
        total_price = float(total_price)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid total price: {total_price!r}")

    try:
        length = float(length_meters)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid pipe length: {length_meters!r}")

    # Calculate total mass for the provided length
    mass = calculate_total_mass(length, diameter, sdr)

    if mass <= 0:
        raise ValueError(f"Mass is invalid: {mass}")

    price_per_kg = total_price / mass

    return price_per_kg


def calculate_length_from_mass(total_mass_kg, diameter, sdr):
    """
    Calculate the pipe length in meters given total mass (kg), diameter, and SDR.

    Args:
        total_mass_kg (float or int): Total mass in kilograms.
        diameter (float or int): Pipe diameter in mm.
        sdr (float or int): SDR value.

    Returns:
        float: Length in meters.

    Raises:
        ValueError: If total_mass_kg cannot be converted to float or mass_per_meter is zero.
        KeyError/ValueError: Propagated from load_weight_table for invalid diameter or SDR.
    """
    # Get mass per meter (kg/m) from weight table
    mass_per_meter = load_weight_table(diameter, sdr)

    # Validate and convert total mass
    try:
        total_mass = float(total_mass_kg)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid total mass: {total_mass_kg!r}")

    if mass_per_meter <= 0:
        raise ValueError(f"Mass per meter is invalid: {mass_per_meter}")

    # Compute and return length in meters
    return total_mass / mass_per_meter

