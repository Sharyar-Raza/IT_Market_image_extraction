import os
import re
import io
import time
import random
import ctypes
import requests
import pandas as pd

from PIL import Image
from urllib.parse import quote, urljoin
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# WINDOWS SLEEP PREVENTION
# ============================================================

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001


def prevent_windows_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        )
        print("Windows sleep prevention: ON")
    except Exception as e:
        print("Sleep prevention error:", e)


def restore_windows_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
        )
        print("Windows sleep prevention: OFF")
    except Exception:
        pass


# ============================================================
# WEBSITE / FILES
# ============================================================

BASE_URL = "https://it-market.com/"

SEARCH_URL = (
    "https://it-market.com/en/search?search={sku}"
)


SCRIPT_DIR = os.path.dirname(
    os.path.abspath(__file__)
)


MASTER_EXCEL = os.path.join(
    SCRIPT_DIR,
    "remaining images.xlsx"
)


OUTPUT_EXCEL = os.path.join(
    SCRIPT_DIR,
    "it_market_images_result.xlsx"
)


IMAGE_ROOT = os.path.join(
    SCRIPT_DIR,
    "images"
)


CHECKPOINT_FILE = os.path.join(
    SCRIPT_DIR,
    "it_market_checkpoint.txt"
)


PROFILE_DIR = os.path.join(
    SCRIPT_DIR,
    "playwright_profile_it_market"
)


PREFERRED_SKU_COLUMNS = [
    "SKU / Part Number",
    "SKU",
    "sku",
    "Part Number",
    "Part number",
    "PartNumber",
    "Model",
    "Model Number",
    "MPN",
]


os.makedirs(
    IMAGE_ROOT,
    exist_ok=True
)


os.makedirs(
    PROFILE_DIR,
    exist_ok=True
)


# ============================================================
# SETTINGS
# ============================================================

PAGE_TIMEOUT = 60000

WAIT_TIME_MS = 12000

PAGE_RETRIES = 3

RETRY_DELAY = 5

SEARCH_WAIT = 2.0

PRODUCT_WAIT = 2.0

SECURITY_MAX_WAIT = 240

SECURITY_CHECK_INTERVAL = 2

IMAGE_TIMEOUT = 25

MIN_WIDTH = 120

MIN_HEIGHT = 80

MIN_BYTES = 1800

SLEEP_BETWEEN_SKUS_MIN = 1.5

SLEEP_BETWEEN_SKUS_MAX = 3.0


LAST_PAGE_ERROR = ""


# ============================================================
# BASIC HELPERS
# ============================================================

def safe_text(
    value
):

    if value is None:
        return ""

    value = str(
        value
    ).strip()

    if value.lower() == "nan":
        return ""

    return value


def normalized_sku(
    value
):

    return re.sub(
        r"[^A-Z0-9]+",
        "",
        safe_text(
            value
        ).upper()
    )


def compare_sku(
    requested_sku,
    website_sku
):

    requested = normalized_sku(
        requested_sku
    )

    actual = normalized_sku(
        website_sku
    )

    return bool(
        requested
        and actual
        and requested == actual
    )


def clean_filename(
    value
):

    value = re.sub(
        r'[\\/:*?"<>|]',
        "_",
        safe_text(
            value
        )
    )

    return re.sub(
        r"\s+",
        "_",
        value
    )


def normalize_url(
    url,
    base=BASE_URL
):

    if not url:
        return ""

    url = str(
        url
    ).strip()

    if url.startswith("//"):
        return (
            "https:"
            + url
        )

    return urljoin(
        base,
        url
    )


def pause():

    time.sleep(
        random.uniform(
            SLEEP_BETWEEN_SKUS_MIN,
            SLEEP_BETWEEN_SKUS_MAX
        )
    )


# ============================================================
# CHECKPOINT
# ============================================================

def save_checkpoint(
    next_index
):

    temp = (
        CHECKPOINT_FILE
        + ".tmp"
    )

    with open(
        temp,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            str(
                next_index
            )
        )

    os.replace(
        temp,
        CHECKPOINT_FILE
    )


def load_checkpoint():

    try:

        with open(
            CHECKPOINT_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            value = (
                f.read()
                .strip()
            )

        return (
            int(
                value
            )
            if value
            else 0
        )

    except Exception:

        return 0


# ============================================================
# EXCEL
# ============================================================

def load_input_excel():

    if not os.path.exists(
        MASTER_EXCEL
    ):

        raise FileNotFoundError(
            f"Excel not found: "
            f"{MASTER_EXCEL}"
        )

    return pd.read_excel(
        MASTER_EXCEL,
        sheet_name=0,
        dtype=str
    )


def find_sku_column(
    df
):

    for column in PREFERRED_SKU_COLUMNS:

        if column in df.columns:

            print(
                "SKU column:",
                column
            )

            return column


    for column in df.columns:

        low = (
            str(
                column
            )
            .strip()
            .lower()
        )

        if (
            "sku" in low
            or "part number" in low
            or "partnumber" in low
            or "model" in low
            or low == "mpn"
        ):

            print(
                "SKU column auto-detected:",
                column
            )

            return column


    if len(
        df.columns
    ):

        print(
            "SKU header not found. "
            "Using first column:",
            df.columns[0]
        )

        return df.columns[0]


    raise RuntimeError(
        "Excel has no columns"
    )


def load_output_excel():

    columns = [
        "SKU",
        "Product URL",
        "Image Path",
    ]


    if os.path.exists(
        OUTPUT_EXCEL
    ):

        try:

            df = pd.read_excel(
                OUTPUT_EXCEL,
                dtype=str
            )

            for column in columns:

                if column not in df.columns:

                    df[column] = ""


            return (
                df[
                    columns
                ]
                .fillna("")
            )

        except Exception:

            pass


    return pd.DataFrame(
        columns=columns
    )


def save_output_excel(
    df
):

    temp = os.path.join(
        SCRIPT_DIR,
        "_it_market_images_temp.xlsx"
    )

    df.to_excel(
        temp,
        index=False
    )

    os.replace(
        temp,
        OUTPUT_EXCEL
    )


# ============================================================
# BROWSER
# ============================================================

def start_browser(
    playwright
):

    try:

        context = (
            playwright
            .chromium
            .launch_persistent_context(
                PROFILE_DIR,
                channel="chrome",
                headless=False,
                no_viewport=True,
                args=[
                    "--start-maximized",
                    "--disable-notifications",
                    "--disable-popup-blocking",
                ],
            )
        )

    except Exception as e:

        print(
            "Installed Chrome launch failed:",
            str(
                e
            )[:200]
        )

        print(
            "Trying Playwright Chromium..."
        )

        context = (
            playwright
            .chromium
            .launch_persistent_context(
                PROFILE_DIR,
                headless=False,
                no_viewport=True,
                args=[
                    "--start-maximized",
                    "--disable-notifications",
                    "--disable-popup-blocking",
                ],
            )
        )


    context.set_default_timeout(
        WAIT_TIME_MS
    )


    context.set_default_navigation_timeout(
        PAGE_TIMEOUT
    )


    page = (
        context.pages[0]
        if context.pages
        else context.new_page()
    )


    return (
        context,
        page
    )


def get_body_text(
    page
):

    try:

        return (
            page
            .locator(
                "body"
            )
            .inner_text(
                timeout=3000
            )
            or ""
        )

    except Exception:

        return ""


# ============================================================
# SECURITY WAIT
# ============================================================

SECURITY_MARKERS = [
    "just a moment",
    "verifying you are human",
    "checking your browser",
    "security verification",
    "please wait while we verify",
    "attention required",
]


def page_is_security(
    page
):

    try:

        title = (
            page.title()
            or ""
        ).lower()


        body = (
            get_body_text(
                page
            )
            .lower()
        )


        current_url = (
            page.url
            or ""
        ).lower()


        if any(
            marker in title
            for marker
            in SECURITY_MARKERS
        ):

            return True


        if any(
            marker in body[:6000]
            for marker
            in SECURITY_MARKERS
        ):

            return True


        if (
            "/cdn-cgi/"
            in current_url
        ):

            return True


        for selector in [
            "#challenge-running",
            "#challenge-stage",
            "#challenge-form",
            "iframe[src*='cloudflare']",
            "iframe[src*='turnstile']",
        ]:

            try:

                if (
                    page
                    .locator(
                        selector
                    )
                    .count()
                    > 0
                ):

                    return True

            except Exception:

                pass


        return False


    except Exception:

        return False


def wait_for_security(
    page
):

    global LAST_PAGE_ERROR


    if not page_is_security(
        page
    ):

        return True


    print(
        f"Security verification detected. "
        f"Waiting up to "
        f"{SECURITY_MAX_WAIT}s..."
    )


    print(
        "Agar verification aaye to browser "
        "me manually complete kar dein."
    )


    started = time.time()


    while (
        time.time()
        - started
        < SECURITY_MAX_WAIT
    ):

        if not page_is_security(
            page
        ):

            LAST_PAGE_ERROR = ""

            print(
                "Verification cleared. Continuing..."
            )

            return True


        time.sleep(
            SECURITY_CHECK_INTERVAL
        )


    LAST_PAGE_ERROR = (
        f"Security verification timeout after "
        f"{SECURITY_MAX_WAIT}s"
    )


    return False


# ============================================================
# OPEN PAGE
# ============================================================

def open_page(
    page,
    url
):

    global LAST_PAGE_ERROR


    LAST_PAGE_ERROR = ""


    for attempt in range(
        1,
        PAGE_RETRIES + 1
    ):

        try:

            print(
                f"Opening "
                f"({attempt}/{PAGE_RETRIES}): "
                f"{url}"
            )


            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT
            )


            try:

                page.locator(
                    "body"
                ).wait_for(
                    state="attached",
                    timeout=WAIT_TIME_MS
                )

            except Exception:

                pass


            if page_is_security(
                page
            ):

                return wait_for_security(
                    page
                )


            LAST_PAGE_ERROR = ""

            return True


        except PlaywrightTimeoutError as e:

            LAST_PAGE_ERROR = (
                "Timeout: "
                + str(
                    e
                )[:300]
            )


        except Exception as e:

            LAST_PAGE_ERROR = (
                "Page error: "
                + str(
                    e
                )[:300]
            )


        print(
            LAST_PAGE_ERROR
        )


        if attempt < PAGE_RETRIES:

            time.sleep(
                RETRY_DELAY
            )


    return False


# ============================================================
# SEARCH RESULT EXACT MATCH
# ============================================================

def parse_brand_and_sku(
    title_text
):

    """
    Example:

    Cisco C1300-24XS

    first word = Cisco
    remaining = C1300-24XS
    """

    title_text = safe_text(
        title_text
    )


    parts = (
        title_text
        .split()
    )


    if len(
        parts
    ) < 2:

        return (
            "",
            ""
        )


    brand = (
        parts[0]
        .strip()
    )


    website_sku = (
        " ".join(
            parts[1:]
        )
        .strip()
    )


    return (
        brand,
        website_sku
    )


def find_first_exact_search_result(
    page,
    requested_sku
):

    requested_norm = normalized_sku(
        requested_sku
    )


    seen = set()


    selectors = [
        "#content-main h2 a[href]",
        "#content-main h3 a[href]",
        "#content-main h4 a[href]",
        "#content-main [class*='product'] h2 a[href]",
        "#content-main [class*='product'] h3 a[href]",
        "#content-main [class*='product'] h4 a[href]",
        "#content-main [class*='product'] a[href]",
        "#content-main article a[href]",
        "#content-main a[href]",
    ]


    for selector in selectors:

        try:

            links = page.locator(
                selector
            )


            count = min(
                links.count(),
                400
            )


            for i in range(
                count
            ):

                try:

                    link = links.nth(
                        i
                    )


                    title_text = (
                        link.inner_text(
                            timeout=1000
                        )
                        or ""
                    ).strip()


                    href = normalize_url(
                        link.get_attribute(
                            "href"
                        ),
                        page.url
                    )


                    if not title_text:

                        continue


                    if not href:

                        continue


                    key = (
                        href,
                        title_text
                    )


                    if key in seen:

                        continue


                    seen.add(
                        key
                    )


                    (
                        brand,
                        website_sku
                    ) = parse_brand_and_sku(
                        title_text
                    )


                    if not brand:

                        continue


                    if not website_sku:

                        continue


                    if (
                        normalized_sku(
                            website_sku
                        )
                        != requested_norm
                    ):

                        continue


                    print(
                        ""
                    )


                    print(
                        "========================================"
                    )


                    print(
                        "FIRST EXACT SEARCH RESULT FOUND"
                    )


                    print(
                        "========================================"
                    )


                    print(
                        "Full title:",
                        title_text
                    )


                    print(
                        "Brand:",
                        brand
                    )


                    print(
                        "Website SKU:",
                        website_sku
                    )


                    print(
                        "Product URL:",
                        href
                    )


                    return {
                        "brand":
                            brand,

                        "website_sku":
                            website_sku,

                        "product_url":
                            href,
                    }


                except Exception:

                    pass


        except Exception:

            pass


    return None


# ============================================================
# IMAGE FILTER
# ============================================================

def probable_product_image_url(
    url
):

    if not url:

        return False


    lower = str(
        url
    ).lower()


    blocked = [
        "watermark",
        "watermarked",
        "favicon",
        "sprite",
        "captcha",
        "placeholder",
        "no-image",
        "no_image",
        "noimage",
        "banner",
        "facebook",
        "instagram",
        "twitter",
        "youtube",
        "linkedin",
        "searching",
        "search-image",
        "reverse-image",
        "reverse_image",
        "lens",
        "loader",
        "loading",
        "rating",
        "badge",
        "buyer-protection",
        "trusted",
        "payment",
        "paypal",
    ]


    if any(
        word in lower
        for word in blocked
    ):

        return False


    clean = (
        lower
        .split("?")[0]
        .split("#")[0]
    )


    if clean.endswith(
        (
            ".svg",
            ".gif",
        )
    ):

        return False


    return True


def element_is_bad_area(
    element
):

    try:

        return bool(
            element.evaluate(
                """
                el => {

                    let node = el;

                    for (
                        let i = 0;
                        i < 12 && node;
                        i++
                    ) {

                        const value = (
                            String(node.className || '')
                            + ' '
                            + String(node.id || '')
                        ).toLowerCase();

                        if (
                            value.includes('recommend') ||
                            value.includes('related') ||
                            value.includes('similar') ||
                            value.includes('upsell') ||
                            value.includes('cross-sell') ||
                            value.includes('crosssell') ||
                            value.includes('recently') ||
                            value.includes('footer') ||
                            value.includes('header')
                        ) {
                            return true;
                        }

                        node =
                            node.parentElement;
                    }

                    return false;
                }
                """
            )
        )

    except Exception:

        return False


# ============================================================
# GET IMAGE URLS
# ============================================================

def image_urls_from_element(
    element,
    product_url
):

    urls = []


    for attr in [
        "src",
        "data-src",
        "data-original",
        "data-lazy-src",
        "data-original-src",
        "data-image",
        "data-full",
        "data-full-image",
        "data-large",
        "data-large-image",
        "data-zoom-image",
        "data-thumb",
        "data-thumbnail",
    ]:

        try:

            value = (
                element.get_attribute(
                    attr
                )
                or ""
            ).strip()


            if value:

                urls.append(
                    normalize_url(
                        value,
                        product_url
                    )
                )


        except Exception:

            pass


    for attr in [
        "srcset",
        "data-srcset",
    ]:

        try:

            value = (
                element.get_attribute(
                    attr
                )
                or ""
            ).strip()


            if not value:

                continue


            entries = []


            for item in value.split(
                ","
            ):

                bits = (
                    item.strip()
                    .split()
                )


                if not bits:

                    continue


                weight = 0


                if len(
                    bits
                ) > 1:

                    number = re.sub(
                        r"[^0-9]",
                        "",
                        bits[1]
                    )


                    if number:

                        weight = int(
                            number
                        )


                entries.append(
                    (
                        weight,
                        bits[0]
                    )
                )


            entries.sort(
                key=lambda x: x[0],
                reverse=True
            )


            for (
                _,
                src
            ) in entries:

                urls.append(
                    normalize_url(
                        src,
                        product_url
                    )
                )


        except Exception:

            pass


    return urls


# ============================================================
# JSON PRODUCT IMAGES
# ============================================================

def extract_json_product_images(
    page,
    product_url
):

    urls = []


    def add(
        value
    ):

        if isinstance(
            value,
            str
        ):

            value = normalize_url(
                value,
                product_url
            )


            if value:

                urls.append(
                    value
                )


        elif isinstance(
            value,
            list
        ):

            for item in value:

                add(
                    item
                )


        elif isinstance(
            value,
            dict
        ):

            for key in [
                "url",
                "contentUrl",
                "thumbnailUrl",
            ]:

                if key in value:

                    add(
                        value[
                            key
                        ]
                    )


    try:

        scripts = page.locator(
            "script[type='application/ld+json']"
        )


        for i in range(
            scripts.count()
        ):

            try:

                raw = (
                    scripts
                    .nth(
                        i
                    )
                    .text_content()
                    or ""
                ).strip()


                if not raw:

                    continue


                data = __import__(
                    "json"
                ).loads(
                    raw
                )


                queue = [
                    data
                ]


                while queue:

                    current = (
                        queue.pop()
                    )


                    if isinstance(
                        current,
                        dict
                    ):

                        obj_type = (
                            current.get(
                                "@type",
                                ""
                            )
                        )


                        if isinstance(
                            obj_type,
                            list
                        ):

                            types = [
                                str(
                                    x
                                ).lower()
                                for x
                                in obj_type
                            ]

                        else:

                            types = [
                                str(
                                    obj_type
                                ).lower()
                            ]


                        if any(
                            x == "product"
                            or x.endswith(
                                "/product"
                            )
                            for x
                            in types
                        ):

                            if (
                                "image"
                                in current
                            ):

                                add(
                                    current[
                                        "image"
                                    ]
                                )


                            if (
                                "images"
                                in current
                            ):

                                add(
                                    current[
                                        "images"
                                    ]
                                )


                        for child in (
                            current.values()
                        ):

                            if isinstance(
                                child,
                                (
                                    dict,
                                    list
                                )
                            ):

                                queue.append(
                                    child
                                )


                    elif isinstance(
                        current,
                        list
                    ):

                        queue.extend(
                            current
                        )


            except Exception:

                pass


    except Exception:

        pass


    return urls


# ============================================================
# EXTRACT OPENED PRODUCT PAGE IMAGES
# ============================================================

def extract_product_images(
    page,
    product_url,
    requested_sku
):

    urls = []


    # ========================================================
    # LAZY LOAD PAGE
    # ========================================================

    try:

        for y in [
            0,
            250,
            500,
            750,
            1000,
            1250,
            1600,
        ]:

            page.evaluate(
                f"window.scrollTo(0,{y})"
            )

            time.sleep(
                0.2
            )


        page.evaluate(
            "window.scrollTo(0,0)"
        )


        time.sleep(
            0.5
        )


    except Exception:

        pass


    # ========================================================
    # PRODUCT / GALLERY IMAGE SELECTORS
    # ========================================================

    selectors = [
        "#content-main .swiper-slide img",
        "#content-main [class*='gallery'] img",
        "#content-main [class*='product-image'] img",
        "#content-main [class*='product_image'] img",
        "#content-main [class*='product'] [class*='image'] img",
        "#content-main [class*='thumbnail'] img",
        "#content-main [class*='thumb'] img",
        "#content-main .carousel img",
        "#content-main img",
    ]


    for selector in selectors:

        try:

            images = page.locator(
                selector
            )


            count = min(
                images.count(),
                150
            )


            for i in range(
                count
            ):

                try:

                    image = images.nth(
                        i
                    )


                    if element_is_bad_area(
                        image
                    ):

                        continue


                    alt = (
                        image.get_attribute(
                            "alt"
                        )
                        or ""
                    ).strip().lower()


                    title = (
                        image.get_attribute(
                            "title"
                        )
                        or ""
                    ).strip().lower()


                    cls = (
                        image.get_attribute(
                            "class"
                        )
                        or ""
                    ).strip().lower()


                    description = (
                        f"{alt} "
                        f"{title} "
                        f"{cls}"
                    )


                    if any(
                        bad in description
                        for bad in [
                            "searching this picture",
                            "search this image",
                            "reverse image",
                            "watermark",
                            "placeholder",
                            "no image",
                            "favicon",
                            "buyer protection",
                            "rating",
                            "badge",
                        ]
                    ):

                        continue


                    if (
                        selector
                        == "#content-main img"
                    ):

                        if (
                            "logo"
                            in description
                            or
                            "icon"
                            in description
                        ):

                            continue


                    element_urls = (
                        image_urls_from_element(
                            image,
                            product_url
                        )
                    )


                    for image_url in element_urls:

                        if not probable_product_image_url(
                            image_url
                        ):

                            continue


                        score = 0


                        low_url = (
                            image_url
                            .lower()
                        )


                        req_norm = normalized_sku(
                            requested_sku
                        )


                        if (
                            req_norm
                            and req_norm
                            in normalized_sku(
                                image_url
                            )
                        ):

                            score += 200


                        if (
                            "product"
                            in low_url
                        ):

                            score += 40


                        if (
                            "gallery"
                            in low_url
                        ):

                            score += 40


                        if (
                            "thumb"
                            in low_url
                        ):

                            score += 20


                        if (
                            req_norm
                            and req_norm
                            in normalized_sku(
                                description
                            )
                        ):

                            score += 100


                        urls.append(
                            (
                                score,
                                image_url
                            )
                        )


                except Exception:

                    pass


        except Exception:

            pass


    # ========================================================
    # JSON-LD PRODUCT IMAGES
    # ========================================================

    for image_url in extract_json_product_images(
        page,
        product_url
    ):

        if not probable_product_image_url(
            image_url
        ):

            continue


        score = 25


        if (
            normalized_sku(
                requested_sku
            )
            in normalized_sku(
                image_url
            )
        ):

            score += 200


        urls.append(
            (
                score,
                image_url
            )
        )


    # ========================================================
    # BEST FIRST
    # ========================================================

    urls.sort(
        key=lambda x: x[0],
        reverse=True
    )


    final_urls = []

    seen = set()


    for (
        _,
        image_url
    ) in urls:

        image_url = (
            str(
                image_url
            )
            .replace(
                "&amp;",
                "&"
            )
            .strip()
        )


        if not image_url.startswith(
            (
                "http://",
                "https://",
            )
        ):

            continue


        if not probable_product_image_url(
            image_url
        ):

            continue


        if image_url in seen:

            continue


        seen.add(
            image_url
        )


        final_urls.append(
            image_url
        )


    print(
        "Product image URLs:",
        len(
            final_urls
        )
    )


    return final_urls


# ============================================================
# REQUEST SESSION
# ============================================================

def create_session(
    context,
    page
):

    session = requests.Session()


    try:

        user_agent = page.evaluate(
            "() => navigator.userAgent"
        )


        if user_agent:

            session.headers.update(
                {
                    "User-Agent":
                        str(
                            user_agent
                        ).strip()
                }
            )


    except Exception:

        session.headers.update(
            {
                "User-Agent":
                    "Mozilla/5.0"
            }
        )


    try:

        for cookie in context.cookies():

            session.cookies.set(
                cookie[
                    "name"
                ],
                cookie[
                    "value"
                ]
            )


    except Exception:

        pass


    return session


# ============================================================
# DOWNLOAD IMAGE
# ============================================================

def download_image(
    session,
    image_url,
    product_url
):

    attempts = [
        {
            "Referer":
                product_url,

            "Accept":
                "image/avif,image/webp,"
                "image/apng,image/*,*/*;q=0.8",
        },

        {
            "Accept":
                "image/avif,image/webp,"
                "image/apng,image/*,*/*;q=0.8",
        },
    ]


    for (
        attempt_number,
        headers
    ) in enumerate(
        attempts,
        start=1
    ):

        try:

            response = session.get(
                image_url,
                headers=headers,
                timeout=IMAGE_TIMEOUT,
                allow_redirects=True
            )


            print(
                f"Image HTTP attempt "
                f"{attempt_number}:",
                response.status_code
            )


            if (
                response.status_code
                != 200
            ):

                continue


            content_type = (
                response.headers.get(
                    "Content-Type",
                    ""
                )
                or ""
            ).lower()


            if (
                "text/html"
                in content_type
                or
                "application/json"
                in content_type
            ):

                continue


            if (
                len(
                    response.content
                )
                < MIN_BYTES
            ):

                continue


            image = Image.open(
                io.BytesIO(
                    response.content
                )
            )


            image.load()


            if (
                image.width
                < MIN_WIDTH
                or
                image.height
                < MIN_HEIGHT
            ):

                continue


            return image


        except Exception as e:

            print(
                "Image download error:",
                str(
                    e
                )[:200]
            )


    return None


# ============================================================
# SAVE IMAGES
# ============================================================

def next_free_image_path(
    sku_folder,
    safe_sku,
    first_image
):

    if first_image:

        base = os.path.join(
            sku_folder,
            (
                f"{safe_sku}"
                f"_ITMarket_thumbnail.jpg"
            )
        )


        if not os.path.exists(
            base
        ):

            return base


    number = 1


    while True:

        candidate = os.path.join(
            sku_folder,
            (
                f"{safe_sku}"
                f"_ITMarket_"
                f"{number:02d}.jpg"
            )
        )


        if not os.path.exists(
            candidate
        ):

            return candidate


        number += 1


def save_product_images(
    session,
    sku,
    product_url,
    image_urls
):

    safe_sku = clean_filename(
        sku
    )


    sku_folder = os.path.join(
        IMAGE_ROOT,
        safe_sku
    )


    os.makedirs(
        sku_folder,
        exist_ok=True
    )


    saved_paths = []


    for image_url in image_urls:

        print(
            ""
        )


        print(
            "Trying product image:"
        )


        print(
            image_url
        )


        image = download_image(
            session,
            image_url,
            product_url
        )


        if image is None:

            continue


        filepath = next_free_image_path(
            sku_folder,
            safe_sku,
            first_image=(
                len(
                    saved_paths
                )
                == 0
            )
        )


        try:

            image.convert(
                "RGB"
            ).save(
                filepath,
                "JPEG",
                quality=95
            )


        except Exception as e:

            print(
                "Image save error:",
                e
            )


            continue


        saved_paths.append(
            filepath
        )


        print(
            "IMAGE SAVED:"
        )


        print(
            filepath
        )


    return saved_paths


# ============================================================
# PROCESS ONE SKU
# ============================================================

def process_sku(
    page,
    context,
    sku
):

    search_url = SEARCH_URL.format(
        sku=quote(
            str(
                sku
            ).strip(),
            safe=""
        )
    )


    print(
        ""
    )


    print(
        "SEARCH:",
        sku
    )


    print(
        "SEARCH URL:",
        search_url
    )


    # ========================================================
    # 1) SEARCH
    # ========================================================

    if not open_page(
        page,
        search_url
    ):

        return (
            "RETRY",
            "",
            []
        )


    time.sleep(
        SEARCH_WAIT
    )


    # ========================================================
    # 2) FIRST EXACT SEARCH RESULT
    # ========================================================

    result = find_first_exact_search_result(
        page,
        sku
    )


    if not result:

        print(
            "Exact result not found."
        )


        return (
            "NOT FOUND",
            "",
            []
        )


    website_sku = (
        result[
            "website_sku"
        ]
    )


    product_url = (
        result[
            "product_url"
        ]
    )


    # ========================================================
    # FINAL EXACT GUARD
    # ========================================================

    if not compare_sku(
        sku,
        website_sku
    ):

        print(
            "Exact SKU guard failed."
        )


        return (
            "NOT FOUND",
            "",
            []
        )


    print(
        ""
    )


    print(
        "EXACT SKU MATCH"
    )


    print(
        "Excel SKU:",
        sku
    )


    print(
        "Website SKU:",
        website_sku
    )


    # ========================================================
    # 3) OPEN EXACT PRODUCT
    # ========================================================

    print(
        "Opening exact product:"
    )


    print(
        product_url
    )


    if not open_page(
        page,
        product_url
    ):

        return (
            "RETRY",
            product_url,
            []
        )


    time.sleep(
        PRODUCT_WAIT
    )


    # ========================================================
    # 4) PRODUCT IMAGES
    # ========================================================

    image_urls = extract_product_images(
        page,
        product_url,
        sku
    )


    if not image_urls:

        print(
            "No product images found."
        )


        return (
            "NO IMAGES",
            product_url,
            []
        )


    # ========================================================
    # 5) DOWNLOAD IMAGES
    # ========================================================

    session = create_session(
        context,
        page
    )


    saved_paths = save_product_images(
        session,
        sku,
        product_url,
        image_urls
    )


    if not saved_paths:

        return (
            "NO IMAGES",
            product_url,
            []
        )


    return (
        "DONE",
        product_url,
        saved_paths
    )


# ============================================================
# MAIN
# ============================================================

def main():

    prevent_windows_sleep()


    context = None

    output_df = None


    try:

        input_df = load_input_excel()


        sku_column = find_sku_column(
            input_df
        )


        output_df = load_output_excel()


        total = len(
            input_df
        )


        checkpoint = min(
            max(
                load_checkpoint(),
                0
            ),
            total
        )


        print(
            ""
        )


        print(
            "================================================"
        )


        print(
            "IT-MARKET EXACT RESULT -> OPEN -> SAVE IMAGES"
        )


        print(
            "================================================"
        )


        print(
            "INPUT EXCEL:",
            MASTER_EXCEL
        )


        print(
            "SKU COLUMN:",
            sku_column
        )


        print(
            "TOTAL SKUs:",
            total
        )


        print(
            "CHECKPOINT:",
            checkpoint
        )


        print(
            "OUTPUT EXCEL:",
            OUTPUT_EXCEL
        )


        print(
            "OUTPUT COLUMNS:"
        )


        print(
            "SKU | Product URL | Image Path"
        )


        print(
            "================================================"
        )


        with sync_playwright() as playwright:

            (
                context,
                page
            ) = start_browser(
                playwright
            )


            for index in range(
                checkpoint,
                total
            ):

                sku = safe_text(
                    input_df.at[
                        index,
                        sku_column
                    ]
                )


                if not sku:

                    save_checkpoint(
                        index + 1
                    )

                    continue


                print(
                    ""
                )


                print(
                    "#" * 60
                )


                print(
                    f"{index + 1}/{total} "
                    f"| SKU: {sku}"
                )


                print(
                    "#" * 60
                )


                try:

                    (
                        status,
                        product_url,
                        saved_paths
                    ) = process_sku(
                        page,
                        context,
                        sku
                    )


                    print(
                        "STATUS:",
                        status
                    )


                    # =================================================
                    # DONE -> ONE ROW PER IMAGE
                    # =================================================

                    if (
                        status
                        == "DONE"
                    ):

                        rows = []


                        for path in saved_paths:

                            rows.append(
                                {
                                    "SKU":
                                        sku,

                                    "Product URL":
                                        product_url,

                                    "Image Path":
                                        path,
                                }
                            )


                        output_df = pd.concat(
                            [
                                output_df,
                                pd.DataFrame(
                                    rows
                                ),
                            ],
                            ignore_index=True
                        )


                    # =================================================
                    # NO IMAGE / NOT FOUND / RETRY
                    # =================================================

                    else:

                        output_df = pd.concat(
                            [
                                output_df,

                                pd.DataFrame(
                                    [
                                        {
                                            "SKU":
                                                sku,

                                            "Product URL":
                                                product_url,

                                            "Image Path":
                                                "",
                                        }
                                    ]
                                ),
                            ],
                            ignore_index=True
                        )


                    save_output_excel(
                        output_df
                    )


                    save_checkpoint(
                        index + 1
                    )


                except Exception as e:

                    print(
                        "SKU ERROR:",
                        str(
                            e
                        )[:500]
                    )


                    save_checkpoint(
                        index + 1
                    )


                pause()


    except KeyboardInterrupt:

        print(
            "Stopped by user. Progress saved."
        )


    except Exception as e:

        print(
            "MAIN ERROR:",
            e
        )


    finally:

        try:

            if output_df is not None:

                save_output_excel(
                    output_df
                )

        except Exception:

            pass


        try:

            if context:

                context.close()

        except Exception:

            pass


        restore_windows_sleep()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()