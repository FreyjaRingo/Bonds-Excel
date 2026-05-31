import pandas as pd
import numpy_financial as npf

from bond_utils import parse_date_value, parse_percent_rate


def to_ql_date(date_value):
    import QuantLib as ql

    dt = parse_date_value(date_value)
    return ql.Date(dt.day, dt.month, dt.year)


def previous_coupon_date(settlement_ql, maturity_ql, frequency):
    import QuantLib as ql

    calendar = ql.NullCalendar()
    tenor = ql.Period(frequency)
    coupon_date = maturity_ql

    while coupon_date > settlement_ql:
        coupon_date = calendar.advance(coupon_date, -tenor, ql.Unadjusted)

    return coupon_date


def calculate_mduration(maturity_date, settlement_date, coupon_rate, yield_rate):
    try:
        import QuantLib as ql

        maturity_dt = parse_date_value(maturity_date)
        settlement_dt = parse_date_value(settlement_date)
        if pd.isna(maturity_dt) or pd.isna(settlement_dt) or maturity_dt <= settlement_dt:
            return None

        settlement_ql = to_ql_date(settlement_dt)
        maturity_ql = to_ql_date(maturity_dt)
        frequency = ql.Semiannual
        effective_ql = previous_coupon_date(settlement_ql, maturity_ql, frequency)

        schedule = ql.Schedule(
            effective_ql,
            maturity_ql,
            ql.Period(frequency),
            ql.NullCalendar(),
            ql.Unadjusted,
            ql.Unadjusted,
            ql.DateGeneration.Backward,
            False,
        )
        day_count = ql.ActualActual(ql.ActualActual.ISMA, schedule)

        ql.Settings.instance().evaluationDate = settlement_ql
        bond = ql.FixedRateBond(0, 100.0, schedule, [coupon_rate], day_count)
        interest_rate = ql.InterestRate(yield_rate, day_count, ql.Compounded, frequency)

        return round(ql.BondFunctions.duration(bond, interest_rate, ql.Duration.Modified, settlement_ql), 4)
    except Exception:
        return None


def calculate_row_mduration(row, maturity_col, settlement_date):
    mat_val = row.get("Maturity") if "Maturity" in row else (row.get(maturity_col) if maturity_col else None)
    kup_val = row.get("kupon %")
    yld_val = row.get("y mbi jual")

    if pd.isna(mat_val) or pd.isna(kup_val) or pd.isna(yld_val):
        return None

    coupon_rate = parse_percent_rate(kup_val)
    yield_rate = parse_percent_rate(yld_val)
    if coupon_rate is None or yield_rate is None:
        return None

    return calculate_mduration(mat_val, settlement_date, coupon_rate, yield_rate)


def calculate_rate_impact(val, cut_rate_input, is_hike=False):
    try:
        if pd.isna(val):
            return None

        rate = parse_percent_rate(val)
        if rate is None:
            return None

        result = rate * (cut_rate_input / 100.0)
        if is_hike:
            result = -result

        return round(result, 6)
    except Exception:
        return None


def calculate_years_to_maturity(row, maturity_col, base_date):
    mat = row.get("Maturity") if "Maturity" in row else row.get(maturity_col)
    if pd.isna(mat):
        return None

    try:
        maturity_date = parse_date_value(mat)
        base_dt = pd.to_datetime(base_date)
        diff = (maturity_date - base_dt.normalize()).days / 365.25
        return round(diff, 4)
    except Exception:
        return None


def calculate_price_pv(row, yield_shift_pct, is_hike=True):
    try:
        y_mbi = parse_percent_rate(row.get("y mbi jual"))
        kupon = parse_percent_rate(row.get("kupon %"))
        ytm_years = row.get("Total Year to Maturity")

        if pd.isna(y_mbi) or pd.isna(kupon) or pd.isna(ytm_years):
            return None

        shift = yield_shift_pct / 100.0
        rate = y_mbi + shift if is_hike else y_mbi - shift
        price = -npf.pv(rate, ytm_years, 100 * kupon, 100, when=0)
        return round(price, 4)
    except Exception:
        return None


def add_simulation_columns(
    df,
    maturity_col,
    settlement_date,
    cut_rate_input,
    base_date_input,
    yield_hike_input,
    yield_cut_input,
):
    df = df.copy()

    df["MDURATION"] = df.apply(
        lambda row: calculate_row_mduration(row, maturity_col, settlement_date),
        axis=1,
    )

    df["Rate Hike"] = df["y mbi jual"].apply(
        lambda val: calculate_rate_impact(val, cut_rate_input, is_hike=True)
    )
    df["Rate Cut"] = df["y mbi jual"].apply(
        lambda val: calculate_rate_impact(val, cut_rate_input, is_hike=False)
    )

    if "mbi_jual" in df.columns:
        mbi_jual_num = pd.to_numeric(df["mbi_jual"], errors="coerce")
        df["Rate Hike Price"] = mbi_jual_num + (df["Rate Hike"] * mbi_jual_num)
        df["Rate Cut Price"] = mbi_jual_num + (df["Rate Cut"] * mbi_jual_num)
    else:
        df["Rate Hike Price"] = None
        df["Rate Cut Price"] = None

    df["Total Year to Maturity"] = df.apply(
        lambda row: calculate_years_to_maturity(row, maturity_col, base_date_input),
        axis=1,
    )
    df["Price if Yield Hike"] = df.apply(
        lambda row: calculate_price_pv(row, yield_hike_input, is_hike=True),
        axis=1,
    )
    df["Price if Yield Cut"] = df.apply(
        lambda row: calculate_price_pv(row, yield_cut_input, is_hike=False),
        axis=1,
    )

    return df
