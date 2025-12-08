"""
Compliance Calculator

Calculates covenant compliance metrics:
- Historic Debt Service Cover Ratio (HDSCR)
- Projected Debt Service Cover Ratio (PDSCR)
- Historic Net Debt Yield (HNDY)
- Projected Net Debt Yield (PNDY)
- Loan to Value (LTV)
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from enum import Enum
from datetime import datetime
from pathlib import Path
import yaml
from loguru import logger


class CovenantStatus(Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    WARNING = "WARNING"  # Within 5% of threshold


@dataclass
class CovenantResult:
    """Result of a single covenant calculation."""
    name: str
    value: float
    threshold: float
    operator: str  # ">=" or "<="
    status: CovenantStatus
    margin: float  # How far from threshold (positive = good)
    calculation_breakdown: Dict[str, float]


@dataclass
class ComplianceResult:
    """Full compliance calculation result."""
    quarter: str
    calculation_date: datetime
    covenants: Dict[str, CovenantResult]
    overall_status: CovenantStatus
    warnings: List[str]


class ComplianceCalculator:
    """
    Calculates covenant compliance from Management Accounts data.
    
    Uses the SFA (Senior Facility Agreement) covenant definitions.
    Loads thresholds from config/covenant_thresholds.yaml.
    """
    
    # Mapping from internal covenant IDs to config file keys
    COVENANT_MAPPING = {
        'hdscr': 'historic_debt_service_cover',
        'pdscr': 'projected_debt_service_cover',
        'hndy': 'historic_net_debt_yield',
        'pndy': 'projected_net_debt_yield',
        'ltv': 'loan_to_value',
    }
    
    # Display names for covenants
    COVENANT_NAMES = {
        'hdscr': 'Historic Debt Service Cover Ratio',
        'pdscr': 'Projected Debt Service Cover Ratio',
        'hndy': 'Historic Net Debt Yield',
        'pndy': 'Projected Net Debt Yield',
        'ltv': 'Loan to Value',
    }
    
    def __init__(self, 
                 config_path: str = "config/covenant_thresholds.yaml",
                 warning_margin: float = None):
        """
        Initialize calculator.
        
        Args:
            config_path: Path to covenant thresholds YAML file
            warning_margin: Percentage margin for warning status (default 5% from config)
        """
        self.config = self._load_config(config_path)
        self.thresholds = self._load_thresholds()
        
        # Get warning margin from config or use default
        if warning_margin is None:
            validation_config = self.config.get('validation', {})
            margin_config = validation_config.get('alert_on_margin_breach', {})
            if margin_config.get('enabled', False):
                self.warning_margin = margin_config.get('margin_percentage', 5) / 100.0
            else:
                self.warning_margin = 0.05  # Default 5%
        else:
            self.warning_margin = warning_margin
    
    def _load_config(self, config_path: str) -> dict:
        """Load configuration from YAML file."""
        path = Path(config_path)
        if not path.exists():
            logger.warning(f"Config file not found: {config_path}, using defaults")
            return {}
        
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    
    def _load_thresholds(self) -> Dict[str, dict]:
        """Load and map thresholds from config."""
        thresholds = {}
        covenants_config = self.config.get('covenants', {})
        
        for covenant_id, config_key in self.COVENANT_MAPPING.items():
            covenant_config = covenants_config.get(config_key, {})
            
            threshold_value = covenant_config.get('threshold')
            operator = covenant_config.get('operator', '>=')
            name = self.COVENANT_NAMES.get(covenant_id, covenant_id)
            
            # Use defaults if not in config
            if threshold_value is None:
                defaults = {
                    'hdscr': {'value': 1.20, 'operator': '>='},
                    'pdscr': {'value': 1.20, 'operator': '>='},
                    'hndy': {'value': 0.062, 'operator': '>='},
                    'pndy': {'value': 0.062, 'operator': '>='},
                    'ltv': {'value': 0.62, 'operator': '<='},
                }
                default = defaults.get(covenant_id, {'value': 0.0, 'operator': '>='})
                threshold_value = default['value']
                operator = default['operator']
            
            thresholds[covenant_id] = {
                'value': threshold_value,
                'operator': operator,
                'name': name,
            }
        
        return thresholds
    
    def calculate(self, 
                  net_operating_income: float,
                  debt_service: float,
                  net_debt: float,
                  property_value: float,
                  projected_noi: float = None,
                  projected_debt_service: float = None,
                  quarter: str = None) -> ComplianceResult:
        """
        Calculate all covenant metrics.
        
        Args:
            net_operating_income: Historic NOI (last 12 months)
            debt_service: Total debt service (interest + principal)
            net_debt: Senior loan - cash
            property_value: Market value of real estate
            projected_noi: Forecasted NOI (next 12 months)
            projected_debt_service: Forecasted debt service
            quarter: Quarter identifier (e.g., "Q3 2025")
            
        Returns:
            ComplianceResult with all covenant calculations
        """
        if projected_noi is None:
            projected_noi = net_operating_income
        if projected_debt_service is None:
            projected_debt_service = debt_service
        
        covenants = {}
        warnings = []
        
        # Calculate Historic DSCR
        hdscr = self._calculate_dscr(net_operating_income, debt_service)
        covenants['hdscr'] = self._evaluate_covenant('hdscr', hdscr, {
            'noi': net_operating_income,
            'debt_service': debt_service
        })
        
        # Calculate Projected DSCR
        pdscr = self._calculate_dscr(projected_noi, projected_debt_service)
        covenants['pdscr'] = self._evaluate_covenant('pdscr', pdscr, {
            'projected_noi': projected_noi,
            'projected_debt_service': projected_debt_service
        })
        
        # Calculate Historic NDY
        hndy = self._calculate_ndy(net_operating_income, net_debt)
        covenants['hndy'] = self._evaluate_covenant('hndy', hndy, {
            'noi': net_operating_income,
            'net_debt': net_debt
        })
        
        # Calculate Projected NDY
        pndy = self._calculate_ndy(projected_noi, net_debt)
        covenants['pndy'] = self._evaluate_covenant('pndy', pndy, {
            'projected_noi': projected_noi,
            'net_debt': net_debt
        })
        
        # Calculate LTV
        ltv = self._calculate_ltv(net_debt, property_value)
        covenants['ltv'] = self._evaluate_covenant('ltv', ltv, {
            'net_debt': net_debt,
            'property_value': property_value
        })
        
        # Determine overall status
        overall_status = CovenantStatus.PASS
        for covenant in covenants.values():
            if covenant.status == CovenantStatus.FAIL:
                overall_status = CovenantStatus.FAIL
                warnings.append(f"FAILED: {covenant.name} = {covenant.value:.2%}")
            elif covenant.status == CovenantStatus.WARNING and overall_status == CovenantStatus.PASS:
                overall_status = CovenantStatus.WARNING
                warnings.append(f"WARNING: {covenant.name} = {covenant.value:.2%} (close to threshold)")
        
        return ComplianceResult(
            quarter=quarter or "Unknown",
            calculation_date=datetime.now(),
            covenants=covenants,
            overall_status=overall_status,
            warnings=warnings
        )
    
    def _calculate_dscr(self, noi: float, debt_service: float) -> float:
        """Calculate Debt Service Cover Ratio."""
        if debt_service == 0:
            return float('inf')
        return noi / debt_service
    
    def _calculate_ndy(self, noi: float, net_debt: float) -> float:
        """Calculate Net Debt Yield."""
        if net_debt == 0:
            return float('inf')
        return noi / net_debt
    
    def _calculate_ltv(self, net_debt: float, property_value: float) -> float:
        """Calculate Loan to Value ratio."""
        if property_value == 0:
            return float('inf')
        return net_debt / property_value
    
    def _evaluate_covenant(self, covenant_id: str, value: float, 
                           breakdown: Dict[str, float]) -> CovenantResult:
        """Evaluate a single covenant against its threshold."""
        config = self.thresholds[covenant_id]
        threshold = config['value']
        operator = config['operator']
        name = config['name']
        
        # Determine pass/fail
        if operator == '>=':
            passed = value >= threshold
            margin = (value - threshold) / threshold
        else:  # '<='
            passed = value <= threshold
            margin = (threshold - value) / threshold
        
        # Determine status
        if passed:
            if margin < self.warning_margin:
                status = CovenantStatus.WARNING
            else:
                status = CovenantStatus.PASS
        else:
            status = CovenantStatus.FAIL
        
        return CovenantResult(
            name=name,
            value=value,
            threshold=threshold,
            operator=operator,
            status=status,
            margin=margin,
            calculation_breakdown=breakdown
        )
    
    def format_result(self, result: ComplianceResult) -> str:
        """Format compliance result as readable text."""
        lines = [
            f"Compliance Certificate - {result.quarter}",
            f"Calculated: {result.calculation_date.strftime('%Y-%m-%d %H:%M')}",
            f"Overall Status: {result.overall_status.value}",
            "",
            "Covenant Results:",
            "-" * 60
        ]
        
        for covenant_id, covenant in result.covenants.items():
            status_symbol = "✓" if covenant.status == CovenantStatus.PASS else "✗" if covenant.status == CovenantStatus.FAIL else "⚠"
            
            if covenant_id == 'ltv':
                formatted_value = f"{covenant.value:.1%}"
                formatted_threshold = f"{covenant.threshold:.0%}"
            else:
                formatted_value = f"{covenant.value:.2%}"
                formatted_threshold = f"{covenant.threshold:.0%}"
            
            lines.append(
                f"{status_symbol} {covenant.name}: {formatted_value} "
                f"({covenant.operator} {formatted_threshold})"
            )
        
        if result.warnings:
            lines.extend(["", "Warnings:", "-" * 60])
            for warning in result.warnings:
                lines.append(f"  • {warning}")
        
        return "\n".join(lines)

