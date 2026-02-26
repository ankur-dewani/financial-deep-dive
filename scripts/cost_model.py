"""
Cost Model Module
Models the transition from current Finance & Accounting cost structure
to a Central Finance team using standardized roles and rates.

Central Finance Roles (from provided spreadsheet):
  SVP of Finance:    $200/hr  ($400,000/yr)
  VP of Finance:     $100/hr  ($200,000/yr)
  Finance Manager:    $50/hr  ($100,000/yr)
  Senior Accountant:  $30/hr   ($60,000/yr)
  Accountant:         $15/hr   ($30,000/yr)
"""

from __future__ import annotations


# Central Finance role definitions (from Central Finance Roles.xlsx)
CENTRAL_ROLES = {
    "SVP of Finance":    {"hourly": 200, "annual": 400_000},
    "VP of Finance":     {"hourly": 100, "annual": 200_000},
    "Finance Manager":   {"hourly":  50, "annual": 100_000},
    "Senior Accountant": {"hourly":  30, "annual":  60_000},
    "Accountant":        {"hourly":  15, "annual":  30_000},
}

# Current F&A employee salaries (from Empl. sheet, 18 employees)
CURRENT_FA_SALARIES = [
    77_000,
    91_300,
    42_500,
    43_424,
    88_000,
    121_000,
    176_000,
    41_800,
    82_500,
    60_500,
    121_000,
    46_658,
    0,         # employee with zero recorded cost
    88_937,
    82_500,
    156_200,
    35_213,
    96_140,
]

# Current F&A non-employee OPEX costs (from OPEX - NEmpl. sheet)
CURRENT_FA_OPEX = {
    "Outsourced Services":   1_426_248,
    "External Contractors":    529_469,
    "Occupancy":               313_683,
    "Hosting":                 225_605,
    "Personnel":               197_362,
    "Marketing":                 2_209,
    "Commissions":              -4_098,
    "T&E/Other":              -319_074,
}


def get_current_fa_cost() -> dict:
    """Calculate total current F&A cost by component."""
    headcount_cost = sum(CURRENT_FA_SALARIES)
    non_hc_cost = sum(CURRENT_FA_OPEX.values())
    total = headcount_cost + non_hc_cost

    return {
        "headcount_count": len(CURRENT_FA_SALARIES),
        "headcount_cost": headcount_cost,
        "non_hc_cost": non_hc_cost,
        "non_hc_breakdown": dict(CURRENT_FA_OPEX),
        "total": total,
        "outsourced_pct": (CURRENT_FA_OPEX["Outsourced Services"] + CURRENT_FA_OPEX["External Contractors"]) / total,
    }


def get_target_fa_model() -> dict:
    """Define the target Central Finance team structure.

    Target Central Finance team structure based on role tiers:
      $150K+        -> VP of Finance (1)
      $85K to $150K -> Finance Manager (2)
      $55K to $85K  -> Senior Accountant (5)
      Under $55K    -> Accountant (10)

    Plus: $200K for statutory audits that cannot be eliminated.
    """
    # Proposed Central Finance team
    roles = [
        {"role": "VP of Finance",     "count": 1, "annual": 200_000},
        {"role": "Finance Manager",   "count": 2, "annual": 100_000},
        {"role": "Senior Accountant", "count": 5, "annual":  60_000},
        {"role": "Accountant",        "count": 10, "annual":  30_000},
    ]

    team_cost = sum(r["count"] * r["annual"] for r in roles)
    statutory_audit = 200_000  # required external audit, cannot eliminate
    total = team_cost + statutory_audit
    total_headcount = sum(r["count"] for r in roles)

    return {
        "roles": roles,
        "team_cost": team_cost,
        "statutory_audit": statutory_audit,
        "total": total,
        "headcount": total_headcount,
    }


def get_employee_role_mapping() -> list[dict]:
    """Map each current F&A employee to a Central Finance role tier by salary band."""
    mapping = []
    for salary in sorted(CURRENT_FA_SALARIES, reverse=True):
        if salary >= 150_000:
            target = "VP of Finance"
        elif salary >= 85_000:
            target = "Finance Manager"
        elif salary >= 55_000:
            target = "Senior Accountant"
        else:
            target = "Accountant"

        mapping.append({
            "current_salary": salary,
            "target_role": target,
            "target_salary": CENTRAL_ROLES[target]["annual"],
        })

    return mapping


def get_savings_summary() -> dict:
    """Calculate savings from moving F&A to Central Finance model."""
    current = get_current_fa_cost()
    target = get_target_fa_model()

    savings = current["total"] - target["total"]
    savings_pct = savings / current["total"]

    return {
        "current_total": current["total"],
        "target_total": target["total"],
        "annual_savings": savings,
        "savings_pct": savings_pct,
        "current_headcount": current["headcount_count"],
        "target_headcount": target["headcount"],
    }


if __name__ == "__main__":
    print("=== Current F&A Cost ===")
    current = get_current_fa_cost()
    print(f"  Headcount: {current['headcount_count']} employees, ${current['headcount_cost']:,.0f}")
    print(f"  Non-HC OPEX: ${current['non_hc_cost']:,.0f}")
    print(f"  Total: ${current['total']:,.0f}")
    print(f"  Outsourced %: {current['outsourced_pct']:.1%}")

    print("\n=== Target Central Finance Model ===")
    target = get_target_fa_model()
    for r in target["roles"]:
        print(f"  {r['count']}x {r['role']}: ${r['count'] * r['annual']:,.0f}")
    print(f"  Statutory Audit: ${target['statutory_audit']:,.0f}")
    print(f"  Total: ${target['total']:,.0f} ({target['headcount']} people)")

    print("\n=== Savings ===")
    summary = get_savings_summary()
    print(f"  Current: ${summary['current_total']:,.0f}")
    print(f"  Target:  ${summary['target_total']:,.0f}")
    print(f"  Savings: ${summary['annual_savings']:,.0f} ({summary['savings_pct']:.0%} reduction)")

    print("\n=== Employee Role Mapping ===")
    for m in get_employee_role_mapping():
        print(f"  ${m['current_salary']:>9,.0f} -> {m['target_role']:<20s} (${m['target_salary']:,.0f})")
