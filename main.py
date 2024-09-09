from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

# Configuration de Selenium
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--no-sandbox")

service = ChromeService(
    executable_path='C:/Users/pc/OneDrive/Bureau/chromedriver-win64/chromedriver-win64/chromedriver.exe')

text_to_keyword = {
    "SOINS DE LA PEAU": "skincare",
    "MAQUILLAGE": "makeup",
    "PARFUMS": "fragrance",
    "BAIN ET CORPS": "bath-body",
    "CHEVEUX": "hair",
    "ACCESSORIES": "accessories",
    "HOMMES": "men",
    "ENFANTS ET BÉBÉS": "kids-baby"
}

base_url = "https://ma.oriflame.com/"

def get_categories(driver):
    urls = []
    try:
        driver.get("https://ma.oriflame.com/")
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="__next"]/main/div[2]/div/div[3]/div'))
        )
        target_div = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[2]/div/div[3]/div')
        buttons = target_div.find_elements(By.TAG_NAME, 'button')

        for button in buttons:
            button_text = button.text
            keyword = text_to_keyword.get(button_text)
            if keyword:
                url = base_url + keyword
                urls.append(url)
    except Exception as e:
        print(f"Erreur lors de la récupération des catégories: {e}")
    return urls


def get_products(driver, url):
    products = []
    try:
        driver.get(url)
        print(f"Navigué vers l'URL: {url}")

        wait = WebDriverWait(driver, 20)
        last_product_count = 0
        clicks = 0
        max_clicks = 20
        delay = 5

        while True:
            try:
                # Essayer de localiser le bouton de chargement avec les IDs possibles
                try:
                    load_more_button = wait.until(
                        EC.element_to_be_clickable((By.ID, ":r8:"))
                    )
                except:
                    load_more_button = wait.until(
                        EC.element_to_be_clickable((By.ID, ":r7:"))
                    )

                # Cliquer sur le bouton
                load_more_button.click()
                print("Bouton 'Charger plus' cliqué")
                clicks += 1

                # Attendre que les nouveaux produits soient chargés
                time.sleep(delay)

                # Vérifier le nombre de produits après le clic
                current_product_count = len(driver.find_elements(By.CSS_SELECTOR, "a.products-app-emotion-azcap3"))
                if current_product_count == last_product_count:
                    print("Aucun nouveau produit trouvé après le clic.")
                    break

                last_product_count = current_product_count

                # Limiter le nombre de clics sur le bouton 'Charger plus'
                if clicks >= max_clicks:
                    print("Nombre maximum de clics atteint.")
                    break

            except Exception as e:
                print("Bouton 'Charger plus' non trouvé ou erreur lors du clic:", e)
                break  # Sortir de la boucle si le bouton ne peut plus être trouvé ou cliqué

        # Récupérer tous les produits après le chargement complet
        try:
            product_links = wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.products-app-emotion-azcap3"))
            )

            for link in product_links:
                href = link.get_attribute("href")
                product_name = link.find_element(By.CSS_SELECTOR,
                                                 "p.MuiTypography-root.MuiTypography-body1.products-app-emotion-9ywzxu").text

                try:
                    product_brand = link.find_element(By.CSS_SELECTOR,
                                                      "span.MuiTypography-root.MuiTypography-caption.products-app-emotion-10g8y54").text
                except:
                    product_brand = "Not Available"

                try:
                    product_price = link.find_element(By.CSS_SELECTOR,
                                                      "p.MuiTypography-root.MuiTypography-body1.products-app-emotion-e7ey7r").text
                except:
                    product_price = "Not Available"

                try:
                    product_rating = link.find_element(By.CSS_SELECTOR,
                                                       "div.base-MuiOriStarRating-root.products-app-emotion-187wtir").get_attribute(
                        "title")
                except:
                    product_rating = "Not Available"

                product_info = {
                    'link': href,
                    'name': product_name,
                    'brand': product_brand,
                    'price': product_price,
                    'rating': product_rating
                }
                products.append(product_info)
                print(f"Product Link: {href}")
                print(f"Product Name: {product_name}")
                print(f"Product Brand: {product_brand}")
                print(f"Product Price: {product_price}")
                print(f"Product Rating: {product_rating}")
                print("-------------")
        except Exception as e:
            print(f"Erreur lors de la récupération des produits: {e}")
    except Exception as e:
        print(f"Erreur lors de la récupération des produits: {str(e)}")
    return products



if __name__ == "__main__":
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("Driver Selenium lancé avec succès.")
        categories = get_categories(driver)
        print(f"Catégories récupérées: {categories}")

        all_products = []

        for category in categories:
            print(f"Traitement de la catégorie: {category}")
            products = get_products(driver, category)
            for product in products:
                all_products.append(product)

        driver.quit()
        print("Driver Selenium fermé avec succès.")

        # Conversion des données en DataFrame
        df = pd.DataFrame(all_products)

        # Enregistrement dans un fichier Excel
        df.to_excel('oriflameproducts.xlsx', index=False, engine='openpyxl')
        print("Données enregistrées avec succès dans oriflame_products.xlsx.")
    except Exception as e:
        print(f"Erreur dans le bloc principal: {e}")