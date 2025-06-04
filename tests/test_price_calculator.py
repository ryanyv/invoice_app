import csv
import os
import pytest

import price_calculator as pc

# Helper to read weight per meter from CSV for expected values
CSV_PATH = os.path.join(os.path.dirname(__file__), os.pardir, 'program files', 'DIN_pivot.csv')


def weight_from_csv(diameter, sdr):
    with open(CSV_PATH, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)
    header = [float(h) for h in rows[0][1:]]
    col_index = header.index(float(sdr)) + 1
    for row in rows[1:]:
        if row and row[0] and float(row[0]) == float(diameter):
            value = row[col_index]
            return float(value)
    raise KeyError('Combination not found')


def test_calculate_total_mass_simple():
    diameter = 90
    sdr = 11
    length = 12
    expected_mass_per_meter = weight_from_csv(diameter, sdr)
    expected_total = expected_mass_per_meter * length
    assert pc.calculate_total_mass(length, diameter, sdr) == pytest.approx(expected_total)


def test_calculate_price_simple():
    diameter = 110
    sdr = 17
    length = 5
    price_per_kg = 2000
    mass_per_meter = weight_from_csv(diameter, sdr)
    expected_total = mass_per_meter * length * price_per_kg
    assert pc.calculate_price(length, diameter, sdr, price_per_kg) == pytest.approx(expected_total)


def test_calculate_length_from_mass_simple():
    diameter = 63
    sdr = 9
    length = 20
    mass_per_meter = weight_from_csv(diameter, sdr)
    total_mass = mass_per_meter * length
    assert pc.calculate_length_from_mass(total_mass, diameter, sdr) == pytest.approx(length)
