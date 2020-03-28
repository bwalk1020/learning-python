import requests
import json
import xlsxwriter

class Product:
  name = ""
  freeTrial = ""
  subscribe = ""
  overview = ""
  isCollection = False
  hasLt = False
  collections = []
  ltProduct = {}

  def __init__(self, name, freeTrial, subscribe, overview, collections, ltCollection):
    self.name = name
    self.freeTrial = freeTrial.split("?")[0].replace("-family", "")
    self.subscribe = subscribe.split("?")[0].replace("-family", "")
    self.overview = overview.split("?")[0].replace("-family", "")
    self.collections = collections

    if collections.__len__() > 0:
       self.isCollection = True

    for product in ltCollection:
        ltString = self.name + " LT"
        if ltString == product["name"]:
           self.ltProduct = product
  
  def setCollection(self, value):
      self.isCollection = value

  def setCollections(self, collections):
      self.collections = collections

  def __str__(self):
      toString = self.name + "\n free-trial-link: " + self.freeTrial + "\n subscribe-link: " + self.subscribe + "\n overview-link: " + self.overview
      if self.ltProduct:
          toString = toString + "\n lt-product: " + self.ltProduct["name"] + "\n lt-product-overview: " + self.ltProduct["overview"]
      return toString

def getProductCollectionsFromAemTags(aemTags):
    collections = []
    
    # for tag in aemTags:
    #     if "/mep" in tag or "/me-products" in tag:
    #        if "Media & Entertainment Collection" not in collections:
    #           collections.append("Media & Entertainment Collection")

    if "personalization-tags:all-products/collection/me-products" in aemTags:
      collections.append("Media & Entertainment Collection")
    if "personalization-tags:all-products/collection/mfg-products" in aemTags:
      collections.append("Product Design & Manufacturing Collection")
    if "personalization-tags:all-products/collection/aec-products" in aemTags:
      collections.append("Architecture, Engineering & Construction Collection")
         

    return collections

def getProductFromSupplementalData(supplementalData, collections, ltCollection):
    productName = ""
    freeTrialLink = ""
    subscribeLink = ""
    overviewLink = ""
    for data in supplementalData:
        if data["key"] == "product-name":
           productName = data["value"]
        elif data["key"] == "free-trial-link":
           freeTrialLink = data["value"]
        elif data["key"] == "subscribe-link":
           subscribeLink = data["value"]
        elif data["key"] == "link-1-url":
           overviewLink = data["value"]
    if  productName != "":
        return Product(productName, freeTrialLink, subscribeLink, overviewLink, collections, ltCollection)
    return None

def getAllLtProducts(content):
    ltCollection = []
    
    for productData in content:
        supplementalData = productData['supplemental-data']
        ltProduct = False
        productName = ""
        overviewLink = ""
        for data in supplementalData:
            if data["key"] == "product-name":
               if "LT" in data["value"]:
                  ltProduct = True
                  productName = data["value"]
            elif data["key"] == "link-1-url":
                 overviewLink = data["value"]
        if ltProduct == True:
            ltCollection.append({"name": productName, "overview": overviewLink})
    return ltCollection

def writeProductsToExcel(products, domain):
    collectionMap = {
      "Architecture, Engineering & Construction Collection":"/collections/architecture-engineering-construction/overview",
      "Product Design & Manufacturing Collection":"/collections/product-design-manufacturing/overview",
      "Media & Entertainment Collection":"/collections/media-entertainment/overview"
    }
    workbook = xlsxwriter.Workbook('SearchTypeAhead-BestBet-eu.xlsx') 
    worksheet = workbook.add_worksheet(domain.replace("/en","-en"))
    row = 0
    col = 0

    worksheet.write(row, col, "Search Term") 
    worksheet.write(row, col + 1, "Type Ahead Result")
    worksheet.write(row, col + 2, "Best Bet URL")

    row += 1
    for product in products: 
        worksheet.write(row, col, product.name) 
        worksheet.write(row, col + 1, product.name)
        worksheet.write(row, col + 2, product.overview)

        if product.freeTrial != "":
           row += 1
           worksheet.write(row, col, "") 
           worksheet.write(row, col + 1, product.name + " free trial")
           worksheet.write(row, col + 2, product.freeTrial) 

        if product.freeTrial != "":
           row += 1
           worksheet.write(row, col, "") 
           worksheet.write(row, col + 1, product.name + " subscribe")
           worksheet.write(row, col + 2, product.subscribe)

        if product.ltProduct:
           row += 1
           worksheet.write(row, col, "") 
           worksheet.write(row, col + 1, product.ltProduct["name"])
           worksheet.write(row, col + 2, product.ltProduct["overview"]) 

        for collection in product.collections:
            row += 1
            worksheet.write(row, col, "") 
            worksheet.write(row, col + 1, collection)
            worksheet.write(row, col + 2, "https://" + domain + collectionMap[collection]) 

        studentEducation = [product.name + " student",product.name + " education","student " + product.name , "education " + product.name]
        row += 1
        worksheet.write(row, col, "") 
        worksheet.write(row, col + 1, "")
        worksheet.write(row, col + 2, "") 
        for se in studentEducation:
            row += 1
            worksheet.write(row, col, se) 
            worksheet.write(row, col + 1, product.name + " student license")
            # if len(product.overview.split("/")) > 4:
            #    response = requests.get(url = "https://" + domain + "/education/free-software/" + product.overview.split("/")[4], params = {})
            #    if response.status_code == 200:
            #       worksheet.write(row, col + 2, "https://" + domain + "/education/free-software/" + product.overview.split("/")[4])
            #    else:
            #       worksheet.write(row, col + 2, "NA")
            # else: 
            #    worksheet.write(row, col + 2, "NA")

            row += 1
            worksheet.write(row, col, "") 
            worksheet.write(row, col + 1, product.name + " free trial")
            worksheet.write(row, col + 2, product.freeTrial)

            row += 1
            worksheet.write(row, col, "") 
            worksheet.write(row, col + 1, product.name + " overview")
            worksheet.write(row, col + 2, product.overview)

            row += 1
            worksheet.write(row, col, "") 
            worksheet.write(row, col + 1, "")
            worksheet.write(row, col + 2, "") 

        row += 1
        worksheet.write(row, col, "") 
        worksheet.write(row, col + 1, "")
        worksheet.write(row, col + 2, "") 
        row += 1

    row += 1
    worksheet.write(row, col, "download") 
    worksheet.write(row, col + 1, "free trials")
    worksheet.write(row, col + 2, "https://" + domain + "/free-trials")         

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "download AutoCAD")
    worksheet.write(row, col + 2, "https://" + domain + "/products/autocad/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "download Revit")
    worksheet.write(row, col + 2, "https://" + domain + "/products/revit/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "download Inventor")
    worksheet.write(row, col + 2, "https://" + domain + "/products/inventor/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "download Fusion 360")
    worksheet.write(row, col + 2, "https://" + domain + "/products/fusion-360/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "download Maya")
    worksheet.write(row, col + 2, "https://" + domain + "/products/maya/free-trial")

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "")
    worksheet.write(row, col + 2, "")

    row += 1
    worksheet.write(row, col, "free") 
    worksheet.write(row, col + 1, "free trials")
    worksheet.write(row, col + 2, "https://" + domain + "/free-trials")         

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "free AutoCAD trial")
    worksheet.write(row, col + 2, "https://" + domain + "/products/autocad/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "free Revit trial")
    worksheet.write(row, col + 2, "https://" + domain + "/products/revit/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "free Inventor trial")
    worksheet.write(row, col + 2, "https://" + domain + "/products/inventor/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "free Fusion 360 trial")
    worksheet.write(row, col + 2, "https://" + domain + "/products/fusion-360/free-trial") 

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "free Maya trial")
    worksheet.write(row, col + 2, "https://" + domain + "/products/maya/free-trial")

    row += 1
    worksheet.write(row, col, "") 
    worksheet.write(row, col + 1, "")
    worksheet.write(row, col + 2, "")



    workbook.close() 
    return

URL = "https://www.autodesk.eu/all-products-data.contenthub.json"
response = requests.get(url = URL, params = {})
allProductsJson  = response.json() 
content = allProductsJson['content']
collectionSet = []
ltCollection = getAllLtProducts(content)
products = []
for item in content:
    collections = []
    supplementalData = item['supplemental-data']
    if "aem-tags" in item:
       collections = getProductCollectionsFromAemTags(item["aem-tags"])
    product = getProductFromSupplementalData(supplementalData, collections, ltCollection)
    if product:
       if product.name and product.freeTrial and product.overview and product.subscribe:
          products.append(product)

products = sorted(products, key=lambda product: product.name)
writeProductsToExcel(products, "www.autodesk.eu")

for product in products:
    print (product)
    if len(product.collections) > 0:
        print (" collections:")
        for collection in product.collections:
            print ("    " + collection)
    print()

