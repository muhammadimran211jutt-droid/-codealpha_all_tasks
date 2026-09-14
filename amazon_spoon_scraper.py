from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import random

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# ==========================================================
# FILE SETTINGS
# ==========================================================

SEARCH_FILE = "Baby_Boy_Toys.html"

OUTPUT_FILE = "Amazon_Baby_Boy_Toys_Details.xlsx"


# ==========================================================
# CHROME SETTINGS
# ==========================================================

options = Options()

# Browser visible rahega
# Agar hidden browser chahiye to is line ko uncomment karein:
# options.add_argument("--headless=new")

options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")
options.add_argument("--lang=en-US")

driver = webdriver.Chrome(options=options)


# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def clean_text(text):

    if not text:
        return ""

    return re.sub(r"\s+", " ", text).strip()


def find_text(soup, selectors):

    for selector in selectors:

        element = soup.select_one(selector)

        if element:

            text = clean_text(
                element.get_text(" ", strip=True)
            )

            if text:
                return text

    return ""


def get_detail_from_table(soup, possible_names):

    """
    Amazon ke Product Details / Technical Details
    table se value find karta hai.
    """

    # Technical details table
    tables = soup.select(
        "#productDetails_techSpec_section_1 tr"
    )

    # Additional details table
    tables += soup.select(
        "#productDetails_detailBullets_sections1 tr"
    )

    for row in tables:

        key = row.select_one("th")
        value = row.select_one("td")

        if not key or not value:
            continue

        key_text = clean_text(
            key.get_text(" ", strip=True)
        )

        value_text = clean_text(
            value.get_text(" ", strip=True)
        )

        for name in possible_names:

            if key_text.lower() == name.lower():

                return value_text

    return ""


def get_asin(soup, url):

    # First URL se ASIN
    match = re.search(
        r"/dp/([A-Z0-9]{10})",
        url
    )

    if match:
        return match.group(1)

    # HTML attributes se ASIN
    for tag in soup.find_all(
        attrs={"data-asin": True}
    ):

        asin = tag.get("data-asin")

        if asin and len(asin) == 10:
            return asin

    # Product details table
    asin = get_detail_from_table(
        soup,
        ["ASIN", "Item model number"]
    )

    if asin:
        return asin

    return ""


# ==========================================================
# STEP 1
# READ SEARCH RESULTS HTML
# ==========================================================

print("\nReading Amazon search page...")

with open(
    SEARCH_FILE,
    "r",
    encoding="utf-8",
    errors="ignore"
) as file:

    html = file.read()


soup = BeautifulSoup(
    html,
    "html.parser"
)


# ==========================================================
# STEP 2
# FIND PRODUCT LINKS
# ==========================================================

product_links = {}

print("Finding product links...")


for a in soup.find_all("a", href=True):

    href = a.get("href", "")

    # Find Amazon /dp/ASIN
    match = re.search(
        r"/dp/([A-Z0-9]{10})",
        href
    )

    if not match:
        continue

    asin = match.group(1)

    # Clean Amazon URL
    url = f"https://www.amazon.com/dp/{asin}"

    product_links[asin] = url


print(
    "Total unique products found:",
    len(product_links)
)


# ==========================================================
# STEP 3
# CREATE LIST FOR DATA
# ==========================================================

all_products = []


# ==========================================================
# STEP 4
# OPEN EVERY PRODUCT
# ==========================================================

for number, (asin, url) in enumerate(
    product_links.items(),
    start=1
):

    print("\n" + "=" * 70)

    print(
        f"PRODUCT {number}/{len(product_links)}"
    )

    print(
        "ASIN:",
        asin
    )

    print("=" * 70)


    try:

        # Open product page
        driver.get(url)

        # Wait for page
        time.sleep(
            random.uniform(3, 5)
        )


        # Scroll page so dynamic sections load
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight * 0.30);"
        )

        time.sleep(1)

        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight * 0.60);"
        )

        time.sleep(1)

        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);"
        )

        time.sleep(2)


        # Get loaded HTML
        product_html = driver.page_source


        # Parse with BeautifulSoup
        product_soup = BeautifulSoup(
            product_html,
            "html.parser"
        )


        # ==================================================
        # 1. PRODUCT TITLE
        # ==================================================

        product_title = find_text(
            product_soup,
            [
                "#productTitle",
                "#title"
            ]
        )


        # ==================================================
        # 2. ASIN
        # ==================================================

        product_asin = get_asin(
            product_soup,
            url
        )

        if not product_asin:
            product_asin = asin


        # ==================================================
        # 3. PRICE
        # ==================================================

        price = find_text(
            product_soup,
            [
                "#corePriceDisplay_desktop_feature_div .a-offscreen",
                "#corePrice_feature_div .a-offscreen",
                "#apex_desktop .a-offscreen",
                "#priceblock_ourprice",
                "#priceblock_dealprice",
                ".a-price .a-offscreen"
            ]
        )


        # ==================================================
        # 4. RATING
        # ==================================================

        rating = find_text(
            product_soup,
            [
                "#acrPopover",
                "span[data-hook='rating-out-of-text']",
                ".reviewCountTextLinkedHistogram"
            ]
        )

        # Agar title attribute mein rating ho
        if not rating:

            rating_tag = product_soup.select_one(
                "#acrPopover"
            )

            if rating_tag:

                rating = clean_text(
                    rating_tag.get("title", "")
                )


        # ==================================================
        # 5. REVIEW COUNT
        # ==================================================

        review_count = find_text(
            product_soup,
            [
                "#acrCustomerReviewLink",
                "span[data-hook='total-review-count']",
                "#averageCustomerReviews a"
            ]
        )


        # ==================================================
        # 6. BRAND
        # ==================================================

        brand = ""

        brand_tag = product_soup.select_one(
            "#bylineInfo"
        )

        if brand_tag:

            brand = clean_text(
                brand_tag.get_text(
                    " ",
                    strip=True
                )
            )

            brand = re.sub(
                r"^Visit the\s+",
                "",
                brand,
                flags=re.I
            )

            brand = re.sub(
                r"\s+Store$",
                "",
                brand,
                flags=re.I
            )


        # ==================================================
        # 7. COLOR
        # ==================================================

        color = get_detail_from_table(
            product_soup,
            [
                "Color"
            ]
        )


        # ==================================================
        # 8. MATERIAL
        # ==================================================

        material = get_detail_from_table(
            product_soup,
            [
                "Material"
            ]
        )


        # ==================================================
        # 9. STYLE
        # ==================================================

        style = get_detail_from_table(
            product_soup,
            [
                "Style"
            ]
        )


        # ==================================================
        # 10. ITEM WEIGHT
        # ==================================================

        item_weight = get_detail_from_table(
            product_soup,
            [
                "Item Weight"
            ]
        )


        # ==================================================
        # 11. PRODUCT DIMENSIONS
        # ==================================================

        product_dimensions = get_detail_from_table(
            product_soup,
            [
                "Product Dimensions"
            ]
        )


        # ==================================================
        # 12. MANUFACTURER
        # ==================================================

        manufacturer = get_detail_from_table(
            product_soup,
            [
                "Manufacturer"
            ]
        )


        # ==================================================
        # 13. MODEL NUMBER
        # ==================================================

        model_number = get_detail_from_table(
            product_soup,
            [
                "Item model number",
                "Model Number"
            ]
        )


        # ==================================================
        # 14. COUNTRY OF ORIGIN
        # ==================================================

        country = get_detail_from_table(
            product_soup,
            [
                "Country of origin",
                "Country of Origin"
            ]
        )


        # ==================================================
        # 15. BEST SELLERS RANK
        # ==================================================

        best_sellers_rank = get_detail_from_table(
            product_soup,
            [
                "Best Sellers Rank"
            ]
        )


        # ==================================================
        # 16. AVAILABILITY
        # ==================================================

        availability = find_text(
            product_soup,
            [
                "#availability",
                "#availability_feature_div"
            ]
        )


        # ==================================================
        # 17. BOUGHT IN PAST MONTH
        # ==================================================

        bought_past_month = find_text(
            product_soup,
            [
                "#social-proofing-faceout",
                "#social-proofing-faceout span",
                ".social-proofing-faceout"
            ]
        )


        # ==================================================
        # 18. PRODUCT DESCRIPTION
        # ==================================================

        description = find_text(
            product_soup,
            [
                "#productDescription",
                "#bookDescription_feature_div"
            ]
        )


        # ==================================================
        # 19. BULLET POINTS / FEATURES
        # ==================================================

        features_list = []

        feature_tags = product_soup.select(
            "#feature-bullets ul li"
        )

        for li in feature_tags:

            text = clean_text(
                li.get_text(
                    " ",
                    strip=True
                )
            )

            if text and text.lower() != "see more":

                features_list.append(text)


        bullet_points = " | ".join(
            features_list
        )


        # ==================================================
        # 20. PRODUCT URL
        # ==================================================

        product_url = url


        # ==================================================
        # CREATE PRODUCT RECORD
        # ==================================================

        product_data = {

            "Product Title":
                product_title,

            "ASIN":
                product_asin,

            "Price":
                price,

            "Rating":
                rating,

            "Review Count":
                review_count,

            "Brand":
                brand,

            "Color":
                color,

            "Material":
                material,

            "Style":
                style,

            "Item Weight":
                item_weight,

            "Product Dimensions":
                product_dimensions,

            "Manufacturer":
                manufacturer,

            "Model Number":
                model_number,

            "Country of Origin":
                country,

            "Best Sellers Rank":
                best_sellers_rank,

            "Availability":
                availability,

            "Bought in past month":
                bought_past_month,

            "Product Description":
                description,

            "Bullet Points / Features":
                bullet_points,

            "Product URL":
                product_url
        }


        # Add to list
        all_products.append(
            product_data
        )


        # ==================================================
        # PRINT RESULT
        # ==================================================

        print(
            "Title:",
            product_title[:80]
        )

        print(
            "Price:",
            price
        )

        print(
            "Rating:",
            rating
        )

        print(
            "Reviews:",
            review_count
        )

        print(
            "Brand:",
            brand
        )

        print(
            "Color:",
            color
        )

        print(
            "Material:",
            material
        )

        print(
            "Style:",
            style
        )

        print(
            "Weight:",
            item_weight
        )


        # ==================================================
        # SAVE AFTER EVERY PRODUCT
        # ==================================================

        df = pd.DataFrame(
            all_products
        )

        df.to_excel(
            OUTPUT_FILE,
            index=False
        )


        print(
            "Saved to Excel."
        )


        # Delay before next product
        time.sleep(
            random.uniform(3, 6)
        )


    except Exception as e:

        print(
            "ERROR:",
            e
        )


        # Even if error occurs,
        # save existing data
        if all_products:

            df = pd.DataFrame(
                all_products
            )

            df.to_excel(
                OUTPUT_FILE,
                index=False
            )


# ==========================================================
# CLOSE BROWSER
# ==========================================================

driver.quit()


# ==========================================================
# FINAL SAVE
# ==========================================================

df = pd.DataFrame(
    all_products
)

df.to_excel(
    OUTPUT_FILE,
    index=False
)


print("\n" + "=" * 70)

print("SCRAPING COMPLETED")

print("=" * 70)

print(
    "Total products scraped:",
    len(all_products)
)

print(
    "Excel file:",
    OUTPUT_FILE
)