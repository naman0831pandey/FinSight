"""
test_price_comparison.py
========================
Automated Test Suite for Smart Suggestion & Price Comparison Engine.

Tests covering:
1. Flipkart is the cheapest below budget.
2. Amazon is the cheapest below budget.
3. Every available offer is above the user's budget.
4. One or more platforms fail to return a valid price (Case D handling).
5. Two platforms have the same lowest price (Tie).
6. A result is for the wrong product variant / accessory rejection.
7. Advanced price normalization (EMIs, discounts, ratings).
8. Semantic relevance & conflicting model/brand rejection.
9. Preserving genuine products (e.g. smartwatches with aluminium case/sport band).
"""

import unittest
from price_comparison_service import (
    ProductOffer,
    evaluate_price_comparison,
    normalize_price,
    is_accessory,
    extract_variant_info,
    compute_relevance_score,
)


class TestPriceComparisonScenarios(unittest.TestCase):

    def setUp(self):
        self.default_budget = 50000.0

    def test_scenario_1_flipkart_cheapest(self):
        """Scenario 1: Flipkart has the lowest price below budget."""
        offers = [
            ProductOffer(platform="Amazon India", product_name="Phone X", price=49500.0, url="https://amazon.in/dp/1"),
            ProductOffer(platform="Flipkart", product_name="Phone X", price=47999.0, url="https://flipkart.com/p/1"),
        ]
        res = evaluate_price_comparison(offers, user_budget=50000.0)
        self.assertEqual(res.case, "A")
        self.assertEqual(len(res.cheapest_offers), 1)
        self.assertEqual(res.cheapest_offers[0].platform, "Flipkart")
        self.assertEqual(res.lowest_verified_price, 47999.0)
        self.assertIn("Flipkart", res.summary_message)
        self.assertFalse(res.has_tie)

    def test_scenario_2_amazon_cheapest(self):
        """Scenario 2: Amazon is cheaper than Flipkart."""
        offers = [
            ProductOffer(platform="Amazon India", product_name="Laptop Pro", price=48999.0, url="https://amazon.in/dp/2"),
            ProductOffer(platform="Flipkart", product_name="Laptop Pro", price=51999.0, url="https://flipkart.com/p/2"),
        ]
        res = evaluate_price_comparison(offers, user_budget=55000.0)
        self.assertEqual(res.case, "A")
        self.assertEqual(len(res.cheapest_offers), 1)
        self.assertEqual(res.cheapest_offers[0].platform, "Amazon India")
        self.assertEqual(res.lowest_verified_price, 48999.0)
        self.assertIn("Amazon India", res.summary_message)
        # Ensure Flipkart was NOT automatically recommended
        self.assertNotIn("Flipkart for ₹48,999", res.summary_message)

    def test_scenario_3_all_offers_above_budget(self):
        """Scenario 3: Every available offer exceeds the user's budget (Case B)."""
        user_budget = 40000.0
        offers = [
            ProductOffer(platform="Amazon India", product_name="Flagship Mobile", price=48999.0, url="https://amazon.in/dp/5"),
            ProductOffer(platform="Flipkart", product_name="Flagship Mobile", price=46999.0, url="https://flipkart.com/p/5"),
        ]
        res = evaluate_price_comparison(offers, user_budget=user_budget)
        self.assertEqual(res.case, "B")
        self.assertEqual(res.lowest_verified_price, 46999.0)
        self.assertEqual(res.cheapest_offers[0].platform, "Flipkart")
        # Explains that it exceeds budget
        self.assertIn("exceeds your budget by ₹6,999.00", res.summary_message)
        self.assertIn("No verified offer is available below your budget", res.summary_message)
        # Check offer budget status labels
        for o in res.offers:
            self.assertIn("Over Budget", o.budget_status)

    def test_scenario_4_platforms_unavailable_or_unverified(self):
        """Scenario 4: One or more platforms fail to return a valid price (Case D handling)."""
        offers = [
            ProductOffer(platform="Amazon India", product_name="Smart Watch", price=None, url="https://amazon.in/s",
                         status="Could not verify (Anti-bot protection)", is_verified=False),
            ProductOffer(platform="Flipkart", product_name="Smart Watch Ultra", price=18999.0, url="https://flipkart.com/p/6",
                         status="Verified Live Price", is_verified=True),
        ]
        res = evaluate_price_comparison(offers, user_budget=20000.0)
        self.assertEqual(res.verified_count, 1)
        self.assertEqual(res.total_count, 2)
        self.assertEqual(res.lowest_verified_price, 18999.0)
        self.assertEqual(res.cheapest_offers[0].platform, "Flipkart")
        # Unverified platforms must have budget_status "Unverified" and price None
        amazon_offer = [o for o in res.offers if o.platform == "Amazon India"][0]
        self.assertIsNone(amazon_offer.price)
        self.assertEqual(amazon_offer.budget_status, "Unverified")

    def test_scenario_5_two_platforms_tie(self):
        """Scenario 5: Both platforms have the exact same lowest price."""
        offers = [
            ProductOffer(platform="Amazon India", product_name="Headphones ANC", price=24990.0, url="https://amazon.in/dp/7"),
            ProductOffer(platform="Flipkart", product_name="Headphones ANC", price=24990.0, url="https://flipkart.com/p/7"),
        ]
        res = evaluate_price_comparison(offers, user_budget=28000.0)
        self.assertTrue(res.has_tie)
        self.assertEqual(len(res.cheapest_offers), 2)
        platforms = {o.platform for o in res.cheapest_offers}
        self.assertEqual(platforms, {"Amazon India", "Flipkart"})
        self.assertIn("tie", res.summary_message.lower())

    def test_scenario_8_variant_and_accessory_filtering(self):
        """Scenario 8: Rejection of accessories and extraction of variant information."""
        # Check accessory detection
        self.assertTrue(is_accessory("Tempered Glass Screen Protector for iPhone 15", "iPhone 15"))
        self.assertTrue(is_accessory("Silicone Back Cover for Samsung S24", "Samsung S24"))
        self.assertTrue(is_accessory("USB-C Fast Charging Cable 2m", "MacBook Air"))
        # If user explicitly searched for cover, should NOT reject
        self.assertFalse(is_accessory("Silicone Cover for iPhone 15", "iPhone 15 cover"))

        # Check variant info extraction
        title_iphone = "Apple iPhone 15 (Blue, 128 GB)"
        self.assertIn("128 GB", extract_variant_info(title_iphone))

        title_tv = "LG 80 cm (32 inch) Full HD Smart LED TV"
        var_tv = extract_variant_info(title_tv)
        self.assertTrue("32 inch" in var_tv or "80 cm" in var_tv)

    def test_price_normalization(self):
        """Test safe price normalization from diverse inputs."""
        self.assertEqual(normalize_price("₹67,999.00"), 67999.0)
        self.assertEqual(normalize_price("Rs. 48,999"), 48999.0)
        self.assertEqual(normalize_price("51,999"), 51999.0)
        self.assertEqual(normalize_price(49999), 49999.0)
        self.assertEqual(normalize_price(""), None)
        self.assertEqual(normalize_price(None), None)
        self.assertEqual(normalize_price("Price on request"), None)

    def test_advanced_price_normalization_discounts_and_emi(self):
        """Ensure EMI rates, star ratings, and savings badges are not mistakenly taken as price."""
        # Ratings with prices
        self.assertEqual(normalize_price("4.5 out of 5 stars ₹49,999"), 49999.0)
        # Savings / discounts
        self.assertEqual(normalize_price("Save ₹2,000. Price: ₹45,999"), 45999.0)
        self.assertEqual(normalize_price("15% off ₹59,999"), 59999.0)
        # EMI rates
        self.assertEqual(normalize_price("₹1,250/month EMI ₹49,999"), 49999.0)

    def test_genuine_product_preservation(self):
        """Ensure real smartwatches or devices containing 'case' / 'band' in official titles are preserved."""
        # Real Apple Watch has Aluminium Case with Sport Band
        title = "Apple Watch Series 9 GPS 45mm Aluminium Case with Sport Band - S/M"
        self.assertFalse(is_accessory(title, "Apple Watch Series 9"))
        self.assertGreaterEqual(compute_relevance_score(title, "Apple Watch Series 9"), 0.8)

        # Real Samsung Galaxy Watch
        title_gw = "Samsung Galaxy Watch6 LTE (44mm, Graphite, Aluminium Case)"
        self.assertFalse(is_accessory(title_gw, "Samsung Galaxy Watch6"))
        self.assertGreaterEqual(compute_relevance_score(title_gw, "Samsung Galaxy Watch6"), 0.8)

    def test_third_party_compatible_and_spare_parts_rejection(self):
        """Ensure third-party compatible accessories and replacement parts are rejected."""
        self.assertTrue(is_accessory("Compatible with iPhone 15 OLED Screen Replacement", "iPhone 15"))
        self.assertTrue(is_accessory("Spigen Ultra Hybrid Back Case Cover for iPhone 15", "iPhone 15"))
        self.assertTrue(is_accessory("Fast Charging Cable for Samsung Galaxy S24", "Samsung Galaxy S24"))
        self.assertTrue(is_accessory("Replacement Ear Cushions for Sony WH-1000XM5", "Sony WH-1000XM5"))
        self.assertTrue(is_accessory("Silicone Strap Band for Apple Watch Series 9", "Apple Watch Series 9"))

    def test_relevance_scoring_conflicting_brands_and_models(self):
        """Ensure conflicting brands and models are rejected (score = 0.0)."""
        # Conflicting models
        self.assertEqual(compute_relevance_score("Apple iPhone 16 128 GB", "iPhone 15"), 0.0)
        self.assertEqual(compute_relevance_score("Sony WH-CH520 Wireless Headphones", "Sony WH-1000XM5"), 0.0)
        self.assertEqual(compute_relevance_score("Galaxy S25 Ultra 5G", "Samsung Galaxy S24"), 0.0)
        # Conflicting brands (sponsored ads)
        self.assertEqual(compute_relevance_score("OnePlus 15 12GB+256GB", "iPhone 15"), 0.0)
        self.assertEqual(compute_relevance_score("POCO X6 5G 8GB RAM 256GB", "iPhone 15"), 0.0)
        # Relevant models match
        self.assertGreaterEqual(compute_relevance_score("Apple iPhone 15 (Blue, 128 GB)", "iPhone 15"), 0.8)
        self.assertGreaterEqual(compute_relevance_score("SONY WH-1000XM5 Wireless Noise Cancellation", "Sony WH-1000XM5"), 0.8)
        self.assertGreaterEqual(compute_relevance_score("boAt Airdopes 141 Elite ANC Bluetooth", "boAt Airdopes 141"), 0.8)


if __name__ == "__main__":
    unittest.main()
