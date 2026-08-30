"""Property categorization shared by the organize script, status checker, and dashboard.

The category folders in the Box root are `0-<Category>`. Since Box cloud rejects
moves of collaborated folders (they revert), most properties still live at the
root — so category is resolved virtually: parent 0-* folder wins, then the name
mapping below, then "Other".
"""

import re
import unicodedata

CATEGORY_FOLDERS = ["0-ABA Care", "0-Dental Care", "0-Retail", "0-Urgent Care"]

MAPPING = {
    "0-Urgent Care": [
        "Access Medical Center (NextCare) - 9917 SE 15th St",
        "Access Medical UrGent Care - 5300 SE 29th St",
        "Access Medical Center - Del City OK",
        "CareNow - 13551 Madison Ave._1",
        "CareNow - 1501 SW Wilshire Blvd._1",
        "CareNow - 155 Wonder World Dr_3",
        "CareNow - 17575 Green Valley Ranch_3",
        "CareNow - 17700 E 39th Street",
        "Christian Family Medicine & Urgent - 79 US-51_2",
        "CityMD - 619 Somerset St._3",
        "Concentra - 1617 N Stoughton Rd",
        "Concentra - 3811 Commons Ave NE",
        "Concentra - 901 W Broadway",
        "Emcura Immediate Care & Primary Care - MedPost - 20599 Mack Ave_3",
        "Fast Pace Health - 2798 Pass Rd_2",
        "Fast Pace Health - 9325 Dayton Pike_5",
        "First Care - 237 N L Rogers Wells Blvd_1",
        "First Care - 60 Fitz Gilbert Rd_1",
        "GoHealth Urgent Care- 10140 Hwy 242_3",
        "MD Now Urgent Care - 901 S State Rd_2",
        "MedExpress - 3397 S Delsea Dr_1",
        "MedFirst Urgent Care - 1315 Chisholm Trail Pkwy_2",
        "MedPost Urgent Care - Nacogdoches, TX",
        "NextCare - 16205 N Pennsylvania Ave_2",
        "NextCare - 603 Texas 35_2",
        "NextCare - 9720 Grant Street_2",
        "Physicians Urgent Care - 406 McBrien Rd",
        "RedMed - 6674 Goodman Rd_1",
        "Xpress Wellness - 2525 Chandler Rd",
        "Xpress Wellness Urgent Care - 2516 N Main St?",
        "Xpress Wellness Urgent Care - 304 S George Nigh Expressway_1",
        "Your Kid's Urgent Care- 4040 49th St N_2",
    ],
    "0-Dental Care": [
        "Aspen Dental - 4063 Mannheim Rd_1",
        "Aspen Dental - Jasper, Indiana",
        "Aspen Dental_Somerset KY",
        "Boise Oral Surgery & Dental Implant Center (Specialty) - 7910 W Ustick Rd",
        "Cornerstone Dentistry - 701 Wilkesboro Blvd Ne_2",
        "Great Lakes Family Dental Group - 9178 US-223",
        "Heartland Dental (9th Avenue Dental Care) - 4850 N 9TH AVE",
        "Heartland Dental (Innovative Dental Care of Muncie)",
        "Heartland Dental (Metro Dental Associates)- 900 52nd St SW",
        "Heartland Dental - 1131 Rutherford Rd_2",
        "Heartland Dental - 2193 Village Mall Dr_5",
        "Heartland Dental - 2886 W Walnut St_1",
        "Heartland Dental - 4520 Lamar Ave_1",
    ],
    "0-Retail": [
        "Applebee's - 1820 W. University Drive_4",
        "Bojangles' - 9375 Dayton Pike",
        "Carl's Jr - 841 SW 89th St_3",
        "CVS - 1301 N Santa Fe Terr",
        "CVS - 1301 North Santa Fe Avenue",
        "CVS - 21 W Main St_1",
        "CVS - 2150 Chester Blvd, Richmond, IN",
        "CVS - 2412 North Classen Boulevard_1",
        "CVS - 3651 West Robinson Street",
        "CVS - 3808 East Washington",
        "CVS - 4000 Battleground Ave",
        "CVS - 5026 US-52, New Palestine, IN",
        "CVS - 5920 Madison Avenue_3",
        "CVS - 789 State Route 39, Martinsville, IN",
        "Del Crest Shops - 4303-4349 SE 15th St_3",
        "Dollar General - 1504 N Yale Ave",
        "Dollar General - 1600 SE 44th St_3",
        "Dollar General - 16841 S Memorial Dr East",
        "Dollar General - 1761 S Old Highway 81_2",
        "Dollar General - 17885 E 116th St. North",
        "Dollar General - 2375 W Omaha St",
        "Dollar General - 324 W Main St_7",
        "Dollar General - 3825 SE 15th St_1",
        "Family Dollar - 3400 N Kelley Ave",
        "Hooters - 13320 N Pennsylvania Ave_1",
        "Masters Car Wash - 2025 NW 142nd St_4",
        "McAlister's Deli",
        "McAlister's Deli - 2006 Eagle Dr_2",
        "Ojos Locos - Phoenix, AZ",
        "Taco Bell - 2403 East Wabash Street_4",
        "Verizon - 2151 Lejeune Blvd_1",
        "Walgreens - 1229 N Eastern Ave_2",
        "Xtreme Auto Wash - 5717 NW 23rd St_2",
    ],
}

# Root folders that are not deals (personal, system, brand-level). Buy-side due-diligence
# folders (DD_*, Due Diligence_*, Buyer_*) ARE deals and stay in.
_EXCLUDE_EXACT = {
    "atlasupdates", "buysidedeals", "fastpace", "fastpacedocumentsfortitle",
    "finalizechecklist", "movingfolder", "myboxnotes", "mycanvases", "sfdc",
    "shared", "urgentcaresales", "jeffgreenman",
}
_EXCLUDE_PREFIXES = ("nathankam", "fantasy", "whitecap", "0")


def squash(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", unicodedata.normalize("NFC", name).casefold())


_CATEGORY_BY_KEY = {
    squash(name): cat.replace("0-", "")
    for cat, names in MAPPING.items()
    for name in names
}


def is_property(folder_name: str) -> bool:
    key = squash(folder_name)
    if key in _EXCLUDE_EXACT:
        return False
    return not any(key.startswith(p) for p in _EXCLUDE_PREFIXES)


def category_of(folder_name: str, parent_name: str = "") -> str:
    if parent_name.startswith("0-"):
        return parent_name.replace("0-", "")
    cat = _CATEGORY_BY_KEY.get(squash(folder_name))
    if cat:
        return cat
    low = folder_name.lower()
    if any(k in low for k in ("urgent care", "immediate care", "nextcare", "fast pace",
                              "access medical", "emcura", "medpost", "carenow")):
        return "Urgent Care"
    if "dental" in low or "dentist" in low:
        return "Dental Care"
    return "Other"
