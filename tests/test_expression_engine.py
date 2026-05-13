import pytest
import numpy as np
from core.expression_engine import ExpressionEngine

class TestExpressionEngine:
    @pytest.fixture
    def engine(self):
        return ExpressionEngine()

    def test_basic_arithmetic(self, engine):
        assert engine.evaluate("10+20", {}) == 30
        assert engine.evaluate("10-5", {}) == 5
        assert engine.evaluate("3*4", {}) == 12
        assert engine.evaluate("10/2", {}) == 5

    def test_with_variables(self, engine):
        context = {"A": 10, "B": 5}
        assert engine.evaluate("A+B", context) == 15
        assert engine.evaluate("A*B", context) == 50

    def test_string_conversion(self, engine):
        context = {"A": "10", "B": " 5 "}
        assert engine.evaluate("A+B", context) == 15

    def test_missing_variable_nan(self, engine):
        engine.missing_value_strategy = 'nan'
        assert np.isnan(engine.evaluate("X+Y", {"X": 10}))

    def test_missing_variable_zero(self):
        engine = ExpressionEngine(missing_value_strategy='zero')
        assert engine.evaluate("X+Y", {"X": 10}) == 10.0

    def test_division_by_zero(self, engine):
        assert np.isnan(engine.evaluate("10/0", {}))

    def test_expression_caching(self, engine):
        expr = "A+B*C"
        context = {"A": 1, "B": 2, "C": 3}
        assert engine.evaluate(expr, context) == 7
        assert expr.upper() in engine._cache

    def test_invalid_expression(self, engine):
        with pytest.raises(ValueError):
            engine.evaluate("10+", {})
