"""
Cross-File Reconciliation Validator

Validates data consistency across multiple files:
- Rent roll total matches executive summary
- Unit sales count matches tracker
- Cash balance matches bank statements
"""

from dataclasses import dataclass, field
from typing import List, Optional
from loguru import logger

from ..parsers.rent_roll_parser import RentRollResult
from ..parsers.sales_tracker_parser import SalesTrackerResult
from ..parsers.bdo_parser import BDOParseResult


@dataclass
class ReconciliationCheck:
    """Result of a single reconciliation check."""
    name: str
    passed: bool
    source_value: float
    target_value: float
    difference: float
    tolerance: float
    message: str
    exact_match: bool = False


@dataclass
class ReconciliationResult:
    """Result of cross-file reconciliation."""
    passed: bool
    checks: List[ReconciliationCheck] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class ReconciliationValidator:
    """
    Validates cross-file data reconciliation.
    
    Checks:
    1. Rent roll total matches executive summary
    2. Unit sales count matches tracker
    3. Cash balance matches bank statements
    """
    
    def __init__(self):
        """Initialize reconciliation validator."""
        pass
    
    def validate(self,
                 rent_roll: Optional[RentRollResult] = None,
                 sales_tracker: Optional[SalesTrackerResult] = None,
                 bdo_result: Optional[BDOParseResult] = None,
                 executive_summary: Optional[dict] = None,
                 management_accounts_cash: Optional[float] = None) -> ReconciliationResult:
        """
        Perform cross-file reconciliation checks.
        
        Args:
            rent_roll: Rent roll parse result
            sales_tracker: Sales tracker parse result
            bdo_result: BDO parse result
            executive_summary: Dict with executive summary values (rent_roll_annual, units_sold)
            management_accounts_cash: Cash value from management accounts
            
        Returns:
            ReconciliationResult with all checks
        """
        checks = []
        warnings = []
        
        # Check 1: Rent roll total
        if rent_roll and executive_summary:
            rent_check = self._reconcile_rent_roll(
                rent_roll.total_annual_rent,
                executive_summary.get('rent_roll_annual', 0.0)
            )
            checks.append(rent_check)
            if not rent_check.passed:
                warnings.append(rent_check.message)
        
        # Check 2: Unit sales count
        if sales_tracker and executive_summary:
            sales_check = self._reconcile_unit_sales(
                sales_tracker.quarter_units_sold,
                executive_summary.get('units_sold', 0)
            )
            checks.append(sales_check)
            if not sales_check.passed:
                warnings.append(sales_check.message)
        
        # Check 3: Cash balance
        if bdo_result and management_accounts_cash is not None:
            cash_check = self._reconcile_cash_balance(
                bdo_result,
                management_accounts_cash
            )
            checks.append(cash_check)
            if not cash_check.passed:
                warnings.append(cash_check.message)
        
        # Overall result
        passed = all(check.passed for check in checks)
        
        return ReconciliationResult(
            passed=passed,
            checks=checks,
            warnings=warnings
        )
    
    def _reconcile_rent_roll(self, 
                            rent_roll_total: float,
                            executive_summary_total: float,
                            tolerance: float = 100.0) -> ReconciliationCheck:
        """
        Reconcile rent roll total with executive summary.
        
        Args:
            rent_roll_total: Total from rent roll parser
            executive_summary_total: Total from executive summary (in thousands)
            tolerance: Tolerance in EUR (default 100)
            
        Returns:
            ReconciliationCheck result
        """
        # Executive summary is typically in thousands, convert to actual
        if executive_summary_total < 10000:  # Likely in thousands
            executive_summary_total = executive_summary_total * 1000
        
        difference = abs(rent_roll_total - executive_summary_total)
        passed = difference <= tolerance
        
        message = (
            f"Rent roll reconciliation: Rent roll (€{rent_roll_total:,.2f}) "
            f"vs Executive Summary (€{executive_summary_total:,.2f}), "
            f"difference: €{difference:,.2f}"
        )
        
        return ReconciliationCheck(
            name="Rent Roll Total",
            passed=passed,
            source_value=rent_roll_total,
            target_value=executive_summary_total,
            difference=difference,
            tolerance=tolerance,
            message=message,
            exact_match=False
        )
    
    def _reconcile_unit_sales(self,
                              sales_tracker_count: int,
                              executive_summary_count: int) -> ReconciliationCheck:
        """
        Reconcile unit sales count.
        
        Args:
            sales_tracker_count: Count from sales tracker
            executive_summary_count: Count from executive summary
            
        Returns:
            ReconciliationCheck result
        """
        difference = abs(sales_tracker_count - executive_summary_count)
        passed = difference == 0  # Exact match required
        
        message = (
            f"Unit sales reconciliation: Sales tracker ({sales_tracker_count} units) "
            f"vs Executive Summary ({executive_summary_count} units), "
            f"difference: {difference} units"
        )
        
        return ReconciliationCheck(
            name="Unit Sales Count",
            passed=passed,
            source_value=float(sales_tracker_count),
            target_value=float(executive_summary_count),
            difference=float(difference),
            tolerance=0.0,
            message=message,
            exact_match=True
        )
    
    def _reconcile_cash_balance(self,
                                bdo_result: BDOParseResult,
                                management_accounts_cash: float,
                                tolerance: float = 0.01) -> ReconciliationCheck:
        """
        Reconcile cash balance between BDO and Management Accounts.
        
        Args:
            bdo_result: BDO parse result
            management_accounts_cash: Cash value from management accounts
            tolerance: Tolerance in EUR (default 0.01)
            
        Returns:
            ReconciliationCheck result
        """
        # Calculate cash total from BDO (11xxxxx accounts)
        bdo_cash_total = 0.0
        for code, entry in bdo_result.accounts.items():
            if code.startswith('11'):  # Cash accounts
                bdo_cash_total += entry.closing_balance
        
        difference = abs(bdo_cash_total - management_accounts_cash)
        passed = difference <= tolerance
        
        message = (
            f"Cash reconciliation: BDO (€{bdo_cash_total:,.2f}) "
            f"vs Management Accounts (€{management_accounts_cash:,.2f}), "
            f"difference: €{difference:,.2f}"
        )
        
        return ReconciliationCheck(
            name="Cash Balance",
            passed=passed,
            source_value=bdo_cash_total,
            target_value=management_accounts_cash,
            difference=difference,
            tolerance=tolerance,
            message=message,
            exact_match=False
        )
    
    def get_summary(self, reconciliation_result: ReconciliationResult) -> str:
        """
        Get human-readable summary of reconciliation results.
        
        Args:
            reconciliation_result: ReconciliationResult
            
        Returns:
            Formatted summary string
        """
        lines = []
        
        if reconciliation_result.passed:
            lines.append("✓ All reconciliations PASSED")
        else:
            lines.append(f"✗ {sum(1 for c in reconciliation_result.checks if not c.passed)} reconciliation(s) FAILED")
        
        for check in reconciliation_result.checks:
            status = "✓" if check.passed else "✗"
            lines.append(f"{status} {check.name}: {check.message}")
        
        if reconciliation_result.warnings:
            lines.append("\nWarnings:")
            for warning in reconciliation_result.warnings:
                lines.append(f"  • {warning}")
        
        return "\n".join(lines)

