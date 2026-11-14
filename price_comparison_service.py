# -*- coding: utf-8 -*-
"""
price_comparison_service.py
===========================
Modular Price Comparison Service for FINSIGHT.

Features:
- Independent search across supported platforms: Amazon India and Flipkart.
- Intelligent semantic query relevance & model matching (eliminates irrelevant products & sponsored ads).
- Advanced accessory & replacement-part filtering (rejects 'compatible with', cases, covers, OLED screens;
  safely preserves genuine products like smartwatches with aluminium cases/sport bands).
- Robust price normalization (excludes EMI rates, discounts, cashback, and rating digits).
- Honest verification status (Verified vs Could Not Verify due to anti-bot / SPA constraints).
- Dynamic best deal selection supporting:
    * Case A: One or more verified offers below user budget (shows lowest verified price + savings).
    * Case B: All verified offers above user budget (shows cheapest available option + excess amount).
    * Case C: Another platform cheaper than Flipkart (dynamic selection based on lowest price).
    * Case D: Unverified/unavailable platforms marked honestly without assuming ₹0 or inventing data.
    * Ties: Identifies equal lowest prices fairly without arbitrarily picking one platform.
- Optional external search provider integration (e.g. SerpApi Google Shopping) via environment variables.
"""

import os
import sys
import re
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple
import concurrent.futures

# Automatically ensure local .venv site-packages is in sys.path
_venv_site = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Lib", "site-packages")
if os.path.exists(_venv_site) and _venv_site not in sys.path:
    sys.path.insert(0, _venv_site)

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False


# Standard modern user-agents with proper client-hints to avoid Akamai/anti-bot interstitials
BROWSER_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-IN,en;q=0.9',
    'DNT': '1',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}

AMAZON_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'en-IN,en;q=0.9,en-US;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'DNT': '1',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Sec-Ch-Ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
    'Sec-Ch-Ua-Mobile': '?0',
    'Sec-Ch-Ua-Platform': '"Windows"',
}

AMAZON_ALT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:128.0) Gecko/20100101 Firefox/128.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'DNT': '1',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
}

FLIPKART_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'Accept-Language': 'en-IN,en;q=0.9',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

# Major known brands and associated product families to prevent cross-brand sponsored ad pollution
KNOWN_BRANDS = {
    'apple': {'iphone', 'ipad', 'macbook', 'airpods'},
    'samsung': {'galaxy'},
    'oneplus': set(),
    'sony': set(),
    'google': {'pixel'},
    'xiaomi': {'redmi', 'poco'},
    'motorola': {'moto'},
    'vivo': {'iqoo'},
    'oppo': {'realme'},
    'boat': set(),
    'noise': set(),
    'dell': set(),
    'hp': set(),
    'lenovo': {'thinkpad'},
    'asus': {'rog'},
    'acer': set()
}

# Explicit accessory phrases that strongly signify an accessory or third-party replacement
ACCESSORY_PHRASES = [
    r'\b(?:compatible\s+with|designed\s+for|suitable\s+for|replacement\s+for|fits?)\b',
    r'\b(?:case\s+cover|back\s+cover|flip\s+cover|silicone\s+case|protective\s+case|phone\s+case|clear\s+case|leather\s+case)\b',
    r'\b(?:tempered\s+glass|screen\s+protector|screen\s+guard|lens\s+protector|camera\s+protector)\b',
    r'\b(?:replacement\s+screen|oled\s+screen|lcd\s+screen|touch\s+screen\s+digitizer|display\s+combo)\b',
    r'\b(?:charging\s+cable|usb\s+cable|type-c\s+cable|lightning\s+cable|fast\s+charging\s+cable)\b',
    r'\b(?:ear\s*tips|ear\s*pads|ear\s*cushions|replacement\s+cushions)\b',
    r'\b(?:skin\s+wrap|sticker\s+skin|body\s+skin)\b',
    r'\b(?:replacement\s+strap|replacement\s+band|watch\s+strap|watch\s+band)\b',
]

# Baseline accessory keywords for query intent detection
ACCESSORY_KEYWORDS = {
    'case', 'cover', 'glass', 'tempered glass', 'screen protector', 'screen guard',
    'cable', 'charger', 'adapter', 'skin', 'strap', 'band', 'pouch', 'sleeve',
    'replacement screen', 'digitizer', 'back cover', 'mount', 'stand', 'remote control',
    'cushion', 'ear tips'
}

PLATFORMS = ["Amazon India", "Flipkart"]


@dataclass
class ProductOffer:
    platform: str
    product_name: str
    price: Optional[float]
    url: str
    availability: str = "Unknown"
    variant: str = ""
    status: str = "Verified"
    is_verified: bool = True
    budget_status: str = "Unverified"
    diff_from_budget: Optional[float] = None
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "platform": self.platform,
            "product_name": self.product_name,
            "price": self.price,
            "url": self.url,
            "availability": self.availability,
            "variant": self.variant,
            "status": self.status,
            "is_verified": self.is_verified,
            "budget_status": self.budget_status,
            "diff_from_budget": self.diff_from_budget,
            "notes": self.notes
        }


@dataclass
class ComparisonResult:
    user_budget: float
    offers: List[ProductOffer]
    cheapest_offers: List[ProductOffer] = field(default_factory=list)
    lowest_verified_price: Optional[float] = None
    case: str = "D"  # 'A', 'B', 'C', 'D'
    summary_message: str = ""
    health_score_impact: int = 0
    checked_time: str = ""
    verified_count: int = 0
    total_count: int = 4
    has_tie: bool = False


# =========================================================================
# --- HELPER FUNCTIONS: NORMALIZATION, RELEVANCE & VARIANT MATCHING ---
# =========================================================================

def normalize_price(price_input: Any) -> Optional[float]:
    """
    Safely normalizes various price formats into a clean numeric float.
    Handles '₹67,999.00', 'Rs. 48,999', '51,999', 48999, None, etc.
    Correctly excludes EMI amounts (e.g. '₹1,250/month'), discount badges
    ('₹2,000 off', '15% off'), and star ratings ('4.5 out of 5 stars').
    """
    if price_input is None:
        return None
    if isinstance(price_input, (int, float)):
        val = float(price_input)
        return val if val > 0 else None
    
    s = str(price_input).strip()
    if not s:
        return None

    # 1. Clean discount / savings prefixes
    s_cleaned = re.sub(r'(?i)\b(?:save|get|upto|up to)\s*(?:[\u20b9]|rs\.?|inr)?\s*[\d,]+(?:\.\d{1,2})?', '', s)
    s_cleaned = re.sub(r'(?i)(?:[\u20b9]|rs\.?|inr)?\s*[\d,]+(?:\.\d{1,2})?\s*(?:%|percent)?\s*(?:off|cashback|discount)', '', s_cleaned)
    # 2. Clean EMI rates
    s_cleaned = re.sub(r'(?i)(?:[\u20b9]|rs\.?|inr)?\s*[\d,]+(?:\.\d{1,2})?\s*\/\s*(?:mo|month|pm)\b', '', s_cleaned)
    s_cleaned = re.sub(r'(?i)\bemi\s*(?:from|starts?\s*at)\s*(?:[\u20b9]|rs\.?|inr)?\s*[\d,]+(?:\.\d{1,2})?', '', s_cleaned)
    s_cleaned = re.sub(r'(?i)\b(?:no\s+cost\s+)?emi\b', '', s_cleaned)
    # 3. Clean rating numbers
    s_cleaned = re.sub(r'(?i)\b\d+(?:\.\d+)?\s*(?:out of|\/5|\/ 5|stars?|ratings?|reviews?)\b', '', s_cleaned)

    # 4. Look for currency symbol followed directly by numeric price
    m = re.search(r'(?:[\u20b9]|rs\.?|inr)\s*([\d,]+(?:\.\d{1,2})?)', s_cleaned, re.IGNORECASE)
    if m:
        try:
            val = float(m.group(1).replace(',', ''))
            return val if val > 0 else None
        except ValueError:
            pass

    # 5. Fallback: Standalone numeric value
    target = s_cleaned if s_cleaned.strip() else s
    target = re.sub(r'(?i)\b(rs|inr|rupees?)\b\.?', '', target)
    target = target.replace('\u20b9', '').replace('₹', '').replace(',', '').strip()
    
    m2 = re.search(r'\b(\d+(?:\.\d{1,2})?)\b', target)
    if m2:
        try:
            val = float(m2.group(1))
            return val if val > 0 else None
        except ValueError:
            return None
            
    return None


def is_accessory(title: str, query: str) -> bool:
    """
    Checks if a product title is an accessory, third-party case, or replacement part
    when the user is looking for a main device.
    
    Preserves genuine products like smartwatches (e.g. 'Apple Watch Series 9 GPS 45mm Aluminium Case with Sport Band')
    while rejecting accessories (e.g. 'Silicone Case for iPhone 15', 'Screen Protector for Samsung S24',
    'Compatible with iPhone 15 OLED Screen Replacement').
    """
    q_lower = query.lower()
    t_lower = title.lower()

    # If the user explicitly searched for an accessory keyword, don't reject
    for kw in ACCESSORY_KEYWORDS:
        if kw in q_lower:
            return False

    # Check clear accessory / compatibility phrases
    for pat in ACCESSORY_PHRASES:
        if re.search(pat, t_lower):
            return True

    # Check for '<accessory_noun> for/of <something>'
    acc_nouns = r'(?:case|cover|glass|guard|protector|strap|band|cable|charger|adapter|battery|screen|digitizer|pouch|sleeve|skin|dock|stand|mount|holder|stylus)'
    if re.search(r'\b' + acc_nouns + r'\s+(?:for|of)\b', t_lower):
        return True

    # Multi-pack accessories like 'Pack of 2', 'Set of 3' at beginning
    if re.search(r'^(?:pack\s+of\s+\d+|set\s+of\s+\d+)\b', t_lower):
        return True

    return False


def extract_model_identifiers(text: str) -> set:
    """
    Extracts alphanumeric model codes (e.g. '15', 's24', '1000xm5', 'm2', '9', 'ps5')
    while filtering out decimal dimensions (e.g. '15.93 cm'), storage units ('128gb'),
    battery sizes ('5000mah'), network generations ('5g'), and years ('2024').
    """
    t = text.lower()
    # Strip decimal dimensions, screen sizes, and unit measurements first
    t = re.sub(r'\b\d+\.\d+\s*(?:cm|inch|-inch|\"|\'|inches)?\b', ' ', t)
    t = re.sub(r'\b\d+\s*(?:cm|inch|-inch|\"|\'|inches|mm|gb|tb|mb|ram|rom|mah|w|hz|mp|g|kg)\b', ' ', t)
    t = re.sub(r'\b(5g|4g|3g|lte|2020|2021|2022|2023|2024|2025|2026|2027)\b', ' ', t)
    tokens = re.findall(r'\b[a-zA-Z0-9]*\d+[a-zA-Z0-9]*\b', t)
    return set(tokens)


def _has_brand(text: str, brand: str) -> bool:
    """Checks for brand presence while avoiding false positives like 'Noise Cancellation'."""
    if brand.lower() == 'noise':
        return bool(re.search(r'\bnoise\b(?!\s*(?:cancellation|cancelling|reduction|canceling|isolation|isolating))', text, re.IGNORECASE))
    return bool(re.search(r'\b' + re.escape(brand) + r'\b', text, re.IGNORECASE))


def compute_relevance_score(title: str, query: str) -> float:
    """
    Computes a relevance score [0.0, 1.0] between a candidate product title and the user's query.
    Enforces brand consistency and model identifier matching, effectively filtering out
    sponsored ads and irrelevant models.
    """
    t_clean = title.lower()
    q_clean = query.lower()

    if is_accessory(title, query):
        return 0.0

    # 1. Brand conflict check: If query specifies a known brand/family, reject other major brands
    # and require presence of the specified brand/family in the candidate title
    for brand, family in KNOWN_BRANDS.items():
        q_has = _has_brand(q_clean, brand) or any(_has_brand(q_clean, f) for f in family)
        if q_has:
            for other_b, other_f in KNOWN_BRANDS.items():
                if other_b != brand:
                    other_in_title = _has_brand(t_clean, other_b) or any(_has_brand(t_clean, f) for f in other_f)
                    other_in_query = _has_brand(q_clean, other_b) or any(_has_brand(q_clean, f) for f in other_f)
                    if other_in_title and not other_in_query:
                        return 0.0
            target_in_title = _has_brand(t_clean, brand) or any(_has_brand(t_clean, f) for f in family)
            if not target_in_title:
                return 0.0

    # 2. Model identifiers check: Query model numbers must be matched in the title
    q_models = extract_model_identifiers(q_clean)
    if q_models:
        t_models = extract_model_identifiers(t_clean)
        # Check direct or substring model match (e.g. 'xm5' in '1000xm5')
        matched = {qm for qm in q_models if any(qm == tm or qm in tm or tm in qm for tm in t_models)}
        if not matched:
            return 0.0
        model_ratio = len(matched) / len(q_models)
    else:
        model_ratio = 1.0

    # 3. Word overlap ratio
    q_words = re.findall(r'[a-zA-Z0-9]+', q_clean)
    t_words = set(re.findall(r'[a-zA-Z0-9]+', t_clean))
    matched_words = [w for w in q_words if w in t_words or any(w in tw for tw in t_words)]
    word_ratio = len(matched_words) / len(q_words) if q_words else 1.0

    # 4. Modifiers alignment ('pro', 'max', 'ultra', 'plus', 'mini', 'air')
    modifiers = {'pro', 'max', 'ultra', 'plus', 'mini', 'air', 'lite'}
    q_mods = {w for w in q_words if w in modifiers}
    t_mods = {w for w in t_words if w in modifiers}
    mod_penalty = 1.0
    if q_mods and not q_mods.issubset(t_mods):
        mod_penalty *= 0.5
    if not q_mods and t_mods:
        mod_penalty *= 0.85

    score = (word_ratio * 0.5 + model_ratio * 0.5) * mod_penalty
    return round(score, 3)


def extract_variant_info(title: str) -> str:
    """
    Extracts key variant details such as storage (128GB, 256GB), screen size (32 inch), or color.
    """
    specs = []
    # Storage / RAM
    storage = re.findall(r'\b\d+\s*(?:GB|TB|MB)\b', title, re.IGNORECASE)
    if storage:
        specs.extend(storage[:2])
    # Screen size
    screen = re.findall(r'\b\d{2,3}(?:\.\d+)?\s*(?:inch|cm|-inch)\b', title, re.IGNORECASE)
    if screen:
        specs.append(screen[0])
    # 5G/4G
    if re.search(r'\b5G\b', title, re.IGNORECASE):
        specs.append("5G")
        
    return ", ".join(specs) if specs else "Standard"


def get_default_platform_search_url(platform: str, query: str) -> str:
    """Returns direct, valid search URL for a given platform."""
    encoded_q = urllib.parse.quote(query.strip())
    if "Amazon" in platform:
        return f"https://www.amazon.in/s?k={encoded_q}"
    elif "Flipkart" in platform:
        return f"https://www.flipkart.com/search?q={encoded_q}"
    return f"https://www.google.com/search?q={encoded_q}"


# =========================================================================
# --- PLATFORM SEARCH WORKERS ---
# =========================================================================

def search_flipkart(query: str, timeout: int = 8) -> ProductOffer:
    """
    Searches Flipkart independently. Retrieves live title, direct product URL, and verified selling price.
    Ranks items by semantic relevance to eliminate accessories and unrelated search results.
    """
    platform = "Flipkart"
    # Sanitize query: strip control characters, limit length to prevent header/parameter injection
    clean_q = re.sub(r'[\r\n\t\x00-\x1f\x7f-\x9f]', ' ', query).strip()
    if len(clean_q) > 100:
        clean_q = clean_q[:100].strip()
    search_url = get_default_platform_search_url(platform, clean_q)
    
    if not REQUESTS_AVAILABLE or not BS4_AVAILABLE:
        return ProductOffer(
            platform=platform,
            product_name=f"{clean_q.title()} on Flipkart",
            price=None,
            url=search_url,
            availability="Check Store",
            status="Could not verify (requests/bs4 not installed)",
            is_verified=False,
            notes="Please install requests and beautifulsoup4."
        )

    try:
        response = requests.get(search_url, headers=FLIPKART_HEADERS, timeout=timeout)
        if response.status_code != 200:
            return ProductOffer(
                platform=platform,
                product_name=f"{clean_q.title()} on Flipkart",
                price=None,
                url=search_url,
                availability="Unknown",
                status=f"Could not verify (HTTP {response.status_code})",
                is_verified=False,
                notes="Search request returned non-200 status code."
            )
            
        soup = BeautifulSoup(response.text, 'html.parser')
        items = soup.select('div[data-id]')
        
        candidates: List[Tuple[float, str, float, str]] = []  # (score, title, price, url)

        for item in items:
            # 1. Extract title (supporting both standard list and grid layout)
            title = ""
            img = item.select_one('img[alt]')
            if img and img.get('alt') and len(img['alt'].strip()) > 5:
                title = img['alt'].strip()
            if not title:
                a_t = item.select_one('a[title]')
                if a_t and a_t.get('title'):
                    title = a_t['title'].strip()
            if not title:
                t_el = item.select_one('div.KzDlHZ, div._4rR01T, a.s1Q9rs, a.wBy4wf, a.atJtCj, div.WKTcLC')
                title = t_el.get_text(strip=True) if t_el else ""
            if not title:
                continue

            # 2. Check relevance
            score = compute_relevance_score(title, clean_q)
            if score < 0.4:
                continue

            # 3. Extract price: prioritize active selling price classes, avoid strikethrough MRP & EMI
            price = None
            p_el = item.select_one('div.Nx9bqj, div._30jeq3, div._16Jk6d')
            if p_el:
                price = normalize_price(p_el.get_text(strip=True))

            if not price:
                for el in item.find_all(True):
                    txt = el.get_text(strip=True)
                    classes = el.get('class', [])
                    # Skip strikethrough MRP classes or discount badges
                    if any(c in classes for c in ['yRaY8j', '_3I9_wc', '_27AhL0']):
                        continue
                    if txt.startswith('₹') and len(el.find_all(True)) == 0 and not any(k in txt.lower() for k in ['off', 'emi', 'month', 'save']):
                        price = normalize_price(txt)
                        if price and price > 0:
                            break

            if not price or price <= 0:
                continue

            # 4. Extract direct product URL
            a = item.select_one('a[href]')
            href = a['href'] if (a and a.get('href')) else ""
            if href.startswith('/'):
                product_url = "https://www.flipkart.com" + href
            elif href.startswith('http'):
                product_url = href
            else:
                product_url = search_url

            candidates.append((score, title, price, product_url))

        # Sort candidates: highest relevance score first, then lowest price
        candidates.sort(key=lambda x: (-x[0], x[2]))

        if candidates:
            best_score, best_title, best_price, best_url = candidates[0]
            variant = extract_variant_info(best_title)
            return ProductOffer(
                platform=platform,
                product_name=best_title,
                price=best_price,
                url=best_url,
                availability="In Stock",
                variant=variant,
                status="Verified Live Price",
                is_verified=True,
                notes="Verified via Flipkart live listing."
            )
                
        # If no relevant item with price was found
        return ProductOffer(
            platform=platform,
            product_name=f"{clean_q.title()} (Search listing)",
            price=None,
            url=search_url,
            availability="Check Store",
            status="Price unavailable",
            is_verified=False,
            notes="No directly matched product with listed price found."
        )
        
    except Exception as e:
        return ProductOffer(
            platform=platform,
            product_name=f"{clean_q.title()} on Flipkart",
            price=None,
            url=search_url,
            availability="Unknown",
            status=f"Could not verify (Network error)",
            is_verified=False,
            notes=str(e)
        )


def search_amazon(query: str, timeout: int = 8) -> ProductOffer:
    """
    Searches Amazon India independently. Retrieves live title, direct product URL, and verified price.
    Ranks items by semantic relevance to eliminate accessories and sponsored ads.
    Gracefully handles bot detection / CAPTCHA without inventing fake prices.
    """
    platform = "Amazon India"
    # Sanitize query: strip control characters, limit length to prevent header/parameter injection
    clean_q = re.sub(r'[\r\n\t\x00-\x1f\x7f-\x9f]', ' ', query).strip()
    if len(clean_q) > 100:
        clean_q = clean_q[:100].strip()
    search_url = get_default_platform_search_url(platform, clean_q)
    
    if not REQUESTS_AVAILABLE or not BS4_AVAILABLE:
        return ProductOffer(
            platform=platform,
            product_name=f"{clean_q.title()} on Amazon",
            price=None,
            url=search_url,
            availability="Check Store",
            status="Could not verify (requests/bs4 not installed)",
            is_verified=False,
            notes="Please install requests and beautifulsoup4."
        )

    # Determine query variants: if searching for iPhone, also consider adding Apple prefix if needed
    query_variants = [clean_q]
    q_low = clean_q.lower()
    if (q_low.startswith("iphone") or q_low.startswith("ipad") or q_low.startswith("macbook")) and not q_low.startswith("apple"):
        query_variants.append(f"Apple {clean_q}")

    headers_pool = [AMAZON_HEADERS, AMAZON_ALT_HEADERS]
    last_error_status = "Could not verify (Network error)"
    anti_bot_detected = False

    for q_attempt in query_variants:
        encoded_term = urllib.parse.quote(q_attempt)
        target_url = f"https://www.amazon.in/s?k={encoded_term}"

        for headers in headers_pool:
            try:
                response = requests.get(target_url, headers=headers, timeout=timeout)
                
                # Check for anti-bot challenges (Akamai bm-verify, AWS WAF, Robot Check, 503)
                if response.status_code == 503 or "bm-verify" in response.text or "Robot Check" in response.text or "api-services-support@amazon.com" in response.text:
                    anti_bot_detected = True
                    last_error_status = "Could not verify (Anti-bot protection)"
                    continue  # Try next headers in pool
                    
                if response.status_code != 200:
                    last_error_status = f"Could not verify (HTTP {response.status_code})"
                    continue

                html = response.content.decode('utf-8', errors='ignore')
                soup = BeautifulSoup(html, 'html.parser')
                cards = soup.find_all('div', attrs={'data-component-type': 's-search-result'})
                
                if not cards:
                    continue

                candidates: List[Tuple[float, str, float, str]] = []  # (score, title, price, url)

                for card in cards:
                    # Collect brand prefix if present (often in a-size-mini header)
                    brand_prefix = ""
                    for h in card.find_all('h2'):
                        h_classes = h.get('class', [])
                        if 'a-size-mini' in h_classes or 's-line-clamp-1' in h_classes:
                            txt = h.get_text(strip=True)
                            if txt and len(txt) < 20:
                                brand_prefix = txt
                                break

                    # 1. Extract title: check aria-label, descriptive h2, a.a-text-normal
                    title = ""
                    for h2 in card.find_all('h2'):
                        if 'a-size-mini' in h2.get('class', []) and len(h2.get_text(strip=True)) < 20:
                            continue
                        if h2.get('aria-label'):
                            title = h2['aria-label'].strip()
                            break
                        t = h2.get_text(strip=True)
                        if len(t) > len(title):
                            title = t

                    if not title or len(title) < 10:
                        for a in card.select('a.a-text-normal, a.s-line-clamp-2, a.s-line-clamp-3, a.s-line-clamp-4'):
                            t = a.get_text(strip=True)
                            if len(t) > len(title):
                                title = t

                    if brand_prefix and brand_prefix.lower() not in title.lower():
                        title = f"{brand_prefix} {title}"

                    if not title or len(title) < 5:
                        continue

                    # 2. Check relevance
                    score = compute_relevance_score(title, clean_q)
                    if score < 0.4:
                        continue

                    # 3. Extract price: prioritize active deal price, avoid strikethrough M.R.P. (a-text-price)
                    p_off = card.select_one('span.a-price:not(.a-text-price) span.a-offscreen') or card.find('span', class_='a-offscreen')
                    p_raw = p_off.get_text(strip=True) if p_off else None
                    price = normalize_price(p_raw)

                    if not price or price <= 0:
                        p_wh = card.select_one('span.a-price:not(.a-text-price) span.a-price-whole') or card.find('span', class_='a-price-whole')
                        price = normalize_price(p_wh.get_text(strip=True) if p_wh else None)

                    if not price or price <= 0:
                        continue

                    # 4. Extract product URL
                    a_link = card.find('a', href=lambda h: h and ('/dp/' in h or '/gp/' in h)) or card.select_one('h2 a')
                    href = a_link['href'] if (a_link and a_link.get('href')) else ""
                    asin = card.get('data-asin')

                    if asin:
                        product_url = f"https://www.amazon.in/dp/{asin}"
                    elif href.startswith('http'):
                        product_url = href
                    elif href:
                        product_url = f"https://www.amazon.in{href}"
                    else:
                        product_url = search_url

                    candidates.append((score, title, price, product_url))

                # Sort candidates: highest relevance score first, then lowest price
                candidates.sort(key=lambda x: (-x[0], x[2]))

                if candidates:
                    best_score, best_title, best_price, best_url = candidates[0]
                    variant = extract_variant_info(best_title)
                    return ProductOffer(
                        platform=platform,
                        product_name=best_title,
                        price=best_price,
                        url=best_url,
                        availability="In Stock",
                        variant=variant,
                        status="Verified Live Price",
                        is_verified=True,
                        notes="Verified via Amazon India live listing."
                    )
            except Exception as e:
                last_error_status = f"Could not verify ({type(e).__name__})"

    if anti_bot_detected:
        return ProductOffer(
            platform=platform,
            product_name=f"{clean_q.title()} on Amazon",
            price=None,
            url=search_url,
            availability="Check Store",
            status="Could not verify (Anti-bot protection)",
            is_verified=False,
            notes="Amazon automated request challenged by anti-bot verification. Click View Deal to check live in browser."
        )

    return ProductOffer(
        platform=platform,
        product_name=f"{clean_q.title()} (Search listing)",
        price=None,
        url=search_url,
        availability="Check Store",
        status="Price unavailable",
        is_verified=False,
        notes="No directly matched product with listed price found."
    )


# =========================================================================
# --- OPTIONAL EXTERNAL API INTEGRATION (SERPAPI / SEARCH PROVIDER) ---
# =========================================================================

def search_via_external_api(query: str) -> Dict[str, ProductOffer]:
    """
    If the user has configured SERPAPI_API_KEY in environment variables,
    queries Google Shopping API to retrieve verified live listings for supported platforms.
    Filters out accessories and irrelevant items.
    Returns a dictionary mapping platform name -> ProductOffer.
    """
    api_key = os.getenv("SERPAPI_API_KEY")
    results: Dict[str, ProductOffer] = {}
    
    if not api_key or not REQUESTS_AVAILABLE:
        return results
        
    try:
        url = "https://serpapi.com/search.json"
        params = {
            "engine": "google_shopping",
            "q": query.strip(),
            "google_domain": "google.co.in",
            "gl": "in",
            "hl": "en",
            "api_key": api_key
        }
        res = requests.get(url, params=params, timeout=10)
        if res.status_code == 200:
            data = res.json()
            shopping_results = data.get("shopping_results", [])
            for item in shopping_results:
                source = item.get("source", "").lower()
                matched_platform = None
                if "amazon" in source:
                    matched_platform = "Amazon India"
                elif "flipkart" in source:
                    matched_platform = "Flipkart"
                    
                if matched_platform and matched_platform not in results:
                    title = item.get("title", "")
                    score = compute_relevance_score(title, query)
                    if score < 0.4:
                        continue

                    price_val = normalize_price(item.get("extracted_price") or item.get("price"))
                    link = item.get("product_link") or item.get("link") or get_default_platform_search_url(matched_platform, query)
                    if price_val and price_val > 0:
                        results[matched_platform] = ProductOffer(
                            platform=matched_platform,
                            product_name=title,
                            price=price_val,
                            url=link,
                            availability="In Stock",
                            variant=extract_variant_info(title),
                            status="Verified Live Price (SerpApi)",
                            is_verified=True,
                            notes="Verified through Google Shopping API."
                        )
    except Exception as e:
        print("External API search error:", e)
        
    return results


# =========================================================================
# --- SEARCH ORCHESTRATOR & COMPARISON EVALUATOR ---
# =========================================================================

def search_all_platforms(query: str) -> List[ProductOffer]:
    """
    Searches supported platforms independently:
    - Amazon India
    - Flipkart
    
    Does NOT hardcode any platform or fabricate prices.
    Uses optional external API if configured, otherwise executes independent workers.
    """
    # Sanitize query: strip control characters, limit length to prevent header/parameter injection
    clean_q = re.sub(r'[\r\n\t\x00-\x1f\x7f-\x9f]', ' ', query).strip()
    if len(clean_q) > 100:
        clean_q = clean_q[:100].strip()
    
    # 1. Check if external API (e.g. SerpApi) is configured
    external_offers = search_via_external_api(clean_q)
    
    # 2. Query each platform concurrently
    offers: List[ProductOffer] = []

    def get_amazon():
        return external_offers.get("Amazon India") or search_amazon(clean_q)

    def get_flipkart():
        return external_offers.get("Flipkart") or search_flipkart(clean_q)

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        f_amz = executor.submit(get_amazon)
        f_flp = executor.submit(get_flipkart)

        # Maintain consistent platform display order: Amazon India, Flipkart
        offers.append(f_amz.result())
        offers.append(f_flp.result())

    return offers


def evaluate_price_comparison(offers: List[ProductOffer], user_budget: float) -> ComparisonResult:
    """
    Dynamically compares collected offers against the user's budget threshold.
    
    Cases implemented:
    - Case A: One or more platforms have a price below the user's budget.
              Recommend the lowest verified price, showing savings.
    - Case B: All verified platforms have prices above the user's budget.
              Clearly state that no verified offer is available below budget.
              Recommend the lowest verified price among available options and
              show how much it exceeds the budget.
    - Case C: Another platform is cheaper than Flipkart (e.g. Amazon India).
              Recommends whichever platform has the genuinely lowest price.
    - Case D: Platform cannot be searched / verified.
              Marked as 'Could not verify' without assuming 0 or inventing prices.
    - Ties:   If multiple platforms have the same lowest price, show the tie.
    """
    now_str = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    verified_offers = [o for o in offers if o.is_verified and o.price is not None and o.price > 0]
    verified_count = len(verified_offers)
    total_count = len(offers)
    
    # Update budget_status and diff_from_budget for each offer
    for o in offers:
        if o.price is not None and o.is_verified:
            diff = o.price - user_budget
            o.diff_from_budget = diff
            if o.price <= user_budget:
                o.budget_status = "Under Budget"
            else:
                o.budget_status = f"Over Budget (+₹{diff:,.2f})"
        else:
            o.budget_status = "Unverified"
            o.diff_from_budget = None

    # Handle Case D (No verified prices available at all)
    if not verified_offers:
        return ComparisonResult(
            user_budget=user_budget,
            offers=offers,
            cheapest_offers=[],
            lowest_verified_price=None,
            case="D",
            summary_message=(
                f"⚠️ Could not verify live prices directly due to platform anti-bot / SPA protections. "
                f"Direct store search links have been provided below so you can inspect prices live in your browser."
            ),
            health_score_impact=0,
            checked_time=now_str,
            verified_count=0,
            total_count=total_count,
            has_tie=False
        )

    # Sort verified offers by price ascending
    verified_offers.sort(key=lambda x: x.price)
    lowest_price = verified_offers[0].price
    
    # Identify all offers matching the lowest price (check for ties)
    cheapest_offers = [o for o in verified_offers if abs(o.price - lowest_price) < 0.01]
    has_tie = len(cheapest_offers) > 1
    
    # Check offers below budget
    under_budget_offers = [o for o in verified_offers if o.price <= user_budget]
    
    if under_budget_offers:
        # CASE A (or Case C if another store is cheaper than Flipkart)
        # Lowest price is <= user_budget
        lowest_under = under_budget_offers[0].price
        cheapest_under = [o for o in under_budget_offers if abs(o.price - lowest_under) < 0.01]
        savings = user_budget - lowest_under
        savings_pct = (savings / user_budget) * 100 if user_budget > 0 else 0
        score_boost = int(min(15, max(1, savings / 1000)))
        
        if len(cheapest_under) > 1:
            names = " & ".join([o.platform for o in cheapest_under])
            msg = (
                f"🎉 Great news! We found a tie deal below your budget!\n"
                f"Both {names} offer this at ₹{lowest_under:,.2f}.\n"
                f"You save ₹{savings:,.2f} ({savings_pct:.1f}% below your budget of ₹{user_budget:,.2f})."
            )
        else:
            best = cheapest_under[0]
            msg = (
                f"🎉 Cheaper Deal Available Below Budget!\n"
                f"Best verified price is on {best.platform} for ₹{best.price:,.2f}.\n"
                f"You save ₹{savings:,.2f} ({savings_pct:.1f}% below your budget of ₹{user_budget:,.2f})."
            )
            
        return ComparisonResult(
            user_budget=user_budget,
            offers=offers,
            cheapest_offers=cheapest_under,
            lowest_verified_price=lowest_under,
            case="A",
            summary_message=msg,
            health_score_impact=score_boost,
            checked_time=now_str,
            verified_count=verified_count,
            total_count=total_count,
            has_tie=(len(cheapest_under) > 1)
        )
    else:
        # CASE B: All verified platforms have prices ABOVE the user's budget
        excess = lowest_price - user_budget
        excess_pct = (excess / user_budget) * 100 if user_budget > 0 else 0
        
        if has_tie:
            names = " & ".join([o.platform for o in cheapest_offers])
            msg = (
                f"ℹ️ No verified offer is available below your budget of ₹{user_budget:,.2f}.\n"
                f"The cheapest available option is a tie between {names} at ₹{lowest_price:,.2f}.\n"
                f"Note: This is the cheapest available option among platforms checked, but it exceeds your budget by ₹{excess:,.2f} (+{excess_pct:.1f}%)."
            )
        else:
            best = cheapest_offers[0]
            msg = (
                f"ℹ️ No verified offer is available below your budget of ₹{user_budget:,.2f}.\n"
                f"The cheapest available option among checked platforms is {best.platform} at ₹{best.price:,.2f}.\n"
                f"Note: This exceeds your budget by ₹{excess:,.2f} (+{excess_pct:.1f}%). We recommend waiting for a sale or revising your budget."
            )
            
        return ComparisonResult(
            user_budget=user_budget,
            offers=offers,
            cheapest_offers=cheapest_offers,
            lowest_verified_price=lowest_price,
            case="B",
            summary_message=msg,
            health_score_impact=0,
            checked_time=now_str,
            verified_count=verified_count,
            total_count=total_count,
            has_tie=has_tie
        )
