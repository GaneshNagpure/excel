from django.db import models

class Instrument(models.Model):
    instrument_name = models.CharField(max_length=255)
    industry_rating = models.CharField(max_length=255)
    quantity = models.IntegerField()
    market_value = models.FloatField()
    percentage_to_nav = models.FloatField()
    isin = models.CharField(max_length=50, unique=True)
    yield_percentage = models.FloatField(null=True, blank=True)
    ytc = models.FloatField(null=True, blank=True)
    instrument_type = models.CharField(max_length=50)  # Equity, Debt, Money Market, etc.
    
    def __str__(self):
        return self.instrument_name

class Portfolio(models.Model):
    name = models.CharField(max_length=255)
    total_market_value = models.FloatField(default=0)
    equity_total = models.FloatField(default=0)
    debt_total = models.FloatField(default=0)
    other_total = models.FloatField(default=0)
    portfolio_date = models.DateField()
    
    def __str__(self):
        return f"Portfolio for {self.name} on {self.portfolio_date}"

class CashEquivalents(models.Model):
    total_value = models.FloatField()
    portfolio = models.ForeignKey(Portfolio, on_delete=models.CASCADE)
    
    def __str__(self):
        return f"Cash Equivalents for {self.portfolio.name}"

