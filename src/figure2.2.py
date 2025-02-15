import os
import pandas as pd
import re
import seaborn as sns
import matplotlib
import matplotlib.pyplot as plt
import numpy as np

FIELD_OF_STUDY_REGEX = re.compile(r"\s*Field\s+of\s+study(1)?(\sand\sgender)?\s*", re.IGNORECASE)
TOTAL_REGEX = re.compile(r"\s*Total(\s+(Awards|Degrees))?\s*", re.IGNORECASE)


def column_include_1(column_name: str) -> bool:
  if FIELD_OF_STUDY_REGEX.match(column_name):
    return True
  if TOTAL_REGEX.match(column_name):
    return True
  return False


read_params_skip3 = {  # type: ignore
  "engine": "calamine",
  "skiprows": 3,
  "usecols": column_include_1,
  "skipfooter": 3,
}

read_params_skip4 = read_params_skip3.copy()
read_params_skip4["skipfooter"] = 4

read_params_skip5 = read_params_skip3.copy()
read_params_skip5["skipfooter"] = 5

read_params_skip9 = read_params_skip3.copy()
read_params_skip9["skipfooter"] = 9

read_params_skip2 = read_params_skip3.copy()
read_params_skip2["skipfooter"] = 2

## DIGEST OF EDUCATION STATISTICS

FILE = os.path.dirname(os.path.realpath(__file__))

## NCES DATA
bachelors1999_00 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors1999-00.xlsx", engine="calamine", skiprows=3, skipfooter=2)
bachelors2000_01 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2000-01.xlsx", **read_params_skip2)
bachelors2001_02 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2001-02.xlsx", **read_params_skip2)
params_override = read_params_skip3.copy()
params_override["skipfooter"] = 6
bachelors2002_03 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2002-03.xlsx", **params_override)
bachelors2003_04 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2003-04.xlsx", **read_params_skip2)
bachelors2004_05 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2004-05.xlsx", **read_params_skip2)
bachelors2004_05.at[1, "Total"] = 1_439_264
bachelors2005_06 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2005-06.xlsx", **read_params_skip2)
bachelors2006_07 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2006-07.xlsx", **read_params_skip3)
bachelors2007_08 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2007-08.xlsx", **read_params_skip5)
bachelors2008_09 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2008-09.xlsx", **read_params_skip9)
bachelors2009_10 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2009-10.xlsx", **read_params_skip5)
bachelors2010_11 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2010-11.xlsx", **read_params_skip3)
bachelors2011_12 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2011-12.xlsx", **read_params_skip4)
# weird data formatting in this xlsx
bachelors2012_13 = pd.read_excel(
  f"{FILE}/../data/fig2.2/raw/Bachelors2012-13.xlsx", engine="calamine", skiprows=4, skipfooter=3, header=None, usecols=(0, 2)
)
bachelors2012_13.rename(columns={0: "Field of study and gender", 2: "Total"}, inplace=True)
bachelors2013_14 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2013-14.xlsx", **read_params_skip3)
## GENERATED SUMMARY TABLES
read_params_for_generated = {  # type: ignore
  "engine": "calamine",
  "skiprows": 4,
  "skipfooter": 4,
}
bachelors2014_15 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2014-15.xlsx", **read_params_for_generated)
bachelors2015_16 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2015-16.xlsx", **read_params_for_generated)
bachelors2016_17 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2016-17.xlsx", **read_params_for_generated)
bachelors2017_18 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2017-18.xlsx", **read_params_for_generated)
bachelors2018_19 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2018-19.xlsx", **read_params_for_generated)
bachelors2019_20 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2019-20.xlsx", **read_params_for_generated)
bachelors2020_21 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2020-21.xlsx", **read_params_for_generated)
bachelors2021_22 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2021-22.xlsx", **read_params_for_generated)
bachelors2022_23 = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors2022-23.xlsx", **read_params_for_generated)


def tidy_method_1(df: pd.DataFrame, year: int) -> pd.DataFrame:
  """
  Used for tidying .xlsx files from 1999-2014, which were documents put together by
  National Center for Education Statistics (NCES), but used IPEDs completion data.
  """

  field_col = None
  total_col = None

  for col in df.columns:
    if FIELD_OF_STUDY_REGEX.match(col):
      field_col = col
    elif TOTAL_REGEX.match(col):
      total_col = col

  if not field_col or not total_col:
    raise ValueError("Field of study or Total column not found.")

  dfcopy = df.copy()

  dfcopy = dfcopy.rename(columns={field_col: "Field_Gender", total_col: "Total"})
  dfcopy = dfcopy.dropna()

  fields = []
  totals = []
  men_totals = []
  women_totals = []

  for i in range(0, len(dfcopy), 3):
    field = dfcopy["Field_Gender"].iloc[i]
    field_total = dfcopy["Total"].iloc[i]

    fields.append(field)
    totals.append(field_total)
    men_totals.append(dfcopy["Total"].iloc[i + 1])
    women_totals.append(dfcopy["Total"].iloc[i + 2])

  tidied_df = pd.DataFrame(
    {
      "Field of Study": fields,
      "Total": totals,
      "Men": men_totals,
      "Women": women_totals,
    }
  )
  tidied_df["Year"] = year

  tidied_df["Percent Men"] = (tidied_df["Men"] / tidied_df["Total"]) * 100
  tidied_df["Percent Women"] = (tidied_df["Women"] / tidied_df["Total"]) * 100

  tidied_df["Percent Men"] = tidied_df["Percent Men"].round(2)
  tidied_df["Percent Women"] = tidied_df["Percent Women"].round(2)

  tidied_df["Total"] = tidied_df["Total"].astype(int)
  tidied_df["Men"] = tidied_df["Men"].astype(int)
  tidied_df["Women"] = tidied_df["Women"].astype(int)

  # sometimes footnotes are parsed into this cell data, remove it.
  tidied_df["Field of Study"] = tidied_df["Field of Study"].str.replace(r"\d+", "", regex=True)

  return tidied_df


def tidy_method_2(df: pd.DataFrame, year: int) -> pd.DataFrame:
  """
  Used for tidying manually generated data using Summary Table on the IPEDs website.
  In this case years from years 2014-23.
  """
  dfcopy = df.copy()
  dfcopy = df.iloc[:, 2:]
  dfcopy.rename(columns={"CIP Title": "Field of Study"}, inplace=True)
  dfcopy["Year"] = year

  dfcopy["Percent Men"] = (dfcopy["Men"] / dfcopy["Total"]) * 100
  dfcopy["Percent Women"] = (dfcopy["Women"] / dfcopy["Total"]) * 100
  dfcopy["Percent Men"] = dfcopy["Percent Men"].round(2)
  dfcopy["Percent Women"] = dfcopy["Percent Women"].round(2)

  dfcopy.at[0, "Field of Study"] = "All fields"

  return dfcopy


all_years_df = pd.concat(
  (
    tidy_method_1(bachelors1999_00, 2000),
    tidy_method_1(bachelors2000_01, 2001),
    tidy_method_1(bachelors2001_02, 2002),
    tidy_method_1(bachelors2002_03, 2003),
    tidy_method_1(bachelors2003_04, 2004),
    tidy_method_1(bachelors2004_05, 2005),
    tidy_method_1(bachelors2005_06, 2006),
    tidy_method_1(bachelors2006_07, 2007),
    tidy_method_1(bachelors2007_08, 2008),
    tidy_method_1(bachelors2008_09, 2009),
    tidy_method_1(bachelors2009_10, 2010),
    tidy_method_1(bachelors2010_11, 2011),
    tidy_method_1(bachelors2011_12, 2012),
    tidy_method_1(bachelors2012_13, 2013),
    tidy_method_1(bachelors2013_14, 2014),
    tidy_method_2(bachelors2014_15, 2015),
    tidy_method_2(bachelors2015_16, 2016),
    tidy_method_2(bachelors2016_17, 2017),
    tidy_method_2(bachelors2017_18, 2018),
    tidy_method_2(bachelors2018_19, 2019),
    tidy_method_2(bachelors2019_20, 2020),
    tidy_method_2(bachelors2020_21, 2021),
    tidy_method_2(bachelors2021_22, 2022),
    tidy_method_2(bachelors2022_23, 2023),
  )
)


def clean_field_of_study(df):
  """Cleans the 'Field of Study' column in a Pandas DataFrame.

  Args:
      df: The Pandas DataFrame.

  Returns:
      The cleaned Pandas DataFrame.
  """

  dfcopy = df.copy()
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.lower()  # Lowercase
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.strip()  # Remove leading/trailing whitespace
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.replace(r"\s+", " ", regex=True)
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.replace(r",\s+and", " and", regex=True)
  # Replace newlines/carriage returns with spaces, then the previous line will turn multiple spaces into single spaces.
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.replace(r"\\xa0", " ", regex=True)
  # Replace non-breaking spaces
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.replace(r"/", " ", regex=True)
  # Replace multiple spaces with single space
  dfcopy["Field of Study"] = dfcopy["Field of Study"].str.replace(r"\s{2,}", " ", regex=True)

  mapping = {
    "agricultural business and production": "agriculture",
    "agricultural sciences": "agriculture",
    "agriculture, agriculture operations": "agriculture",
    "agriculture, agriculture operations, and related sciences": "agriculture",
    "agricultural/animal/plant/veterinary science": "agriculture",
    "architecture and related programs": "architecture",
    "architecture and related services": "architecture",
    "area, ethnic and cultural studies": "area, ethnic, cultural studies",
    "area, ethnic, cultural, and gender studies": "area, ethnic, cultural studies",
    "area, ethnic, cultural, gender, and group studies": "area, ethnic, cultural studies",
    "biological sciences life sciences": "biological sciences",
    "biological and biomedical sciences": "biological sciences",
    "business management and administrative services": "business",
    "business, management, marketing": "business",
    "communication, journalism": "communications",
    "communications technologies": "communications technology",
    "communications technologies technicians": "communications technology",
    "computer and information sciences": "computer science",
    "computer and information sciences and support services": "computer science",
    "conservation and renewable natural resources": "natural resources",
    "natural resources and conservation": "natural resources",
    "engineering technologies technicians": "engineering technology",
    "engineering technologies": "engineering technology",
    "english language and literature letters": "english language and literature",
    "family and consumer sciences human sciences": "family and consumer sciences",
    "foreign languages and literatures": "foreign languages",
    "foreign languages, literatures, and linguistics": "foreign languages",
    "health professions and related sciences": "health professions",
    "health professions and related clinical sciences": "health professions",
    "homeland security, law enforcement, firefighting": "civil service",
    "law and legal studies": "law",
    "legal professions and studies": "law",
    "liberal general studies and humanities": "liberal arts",
    "liberal arts and sciences, general studies, and humanities": "liberal arts",
    "library science": "library science",
    "mathematics and statistics": "mathematics",
    "mechanic and repair technologies technicians": "mechanic and repair",
    "military technologies": "military technology",
    "multi interdisciplinary studies": "multi/interdisciplinary studies",
    "parks, recreation, leisure and fitness": "parks and recreation",
    "parks, recreation, leisure and fitness studies": "parks and recreation",
    "parks, recreation, leisure, fitness, and kinesiology": "parks and recreation",
    "personal and culinary services": "personal services",
    "philosophy and religion": "philosophy and religious studies",
    "physical sciences": "physical sciences",
    "precision production trades": "precision production",
    "protective services": "protective services",
    "public administration and services": "public administration",
    "public administration and social service professions": "public administration",
    "reserve officer training corps (jrotc, rotc)": "military",
    "residency programs": "medical residency",
    "science technologies": "science technology",
    "science technologies technicians": "science technology",
    "security and protective services": "protective services",
    "social sciences and history": "social sciences",
    "technology education industrial arts": "technology education",
    "theological studies and religious vocations": "theology",
    "theology and religious vocations": "theology",
    "transportation and materials moving workers": "transportation",
    "visual and performing arts": "visual and performing arts",
    "vocational home economics": "home economics",
    "other, not specified above": "other",
    "other": "other",
    "all fields": "all fields",
    "construction trades": "construction",
    "culinary, entertainment, and personal services": "personal services",
    "engineering engineering technicians": "engineering",
    "homeland security, law enforcement, firefighting and related protective services": "civil service",
    # overlaps rows...
    "engineering technologies and engineering-related fields": "engineering",
    "agriculture, agriculture operations and related sciences": "agriculture",
    "engineering engineering-related technologies technicians": "engineering",
    "engineering-related technologies": "engineering",
    "engineering technology": "engineering",
  }

  dfcopy["Field of Study"] = dfcopy["Field of Study"].replace(mapping)

  return dfcopy


all_years_df = clean_field_of_study(all_years_df)
all_years_df = all_years_df.drop(columns=["CIP Code"])
all_years_df = all_years_df.reset_index()
all_years_df.to_csv(f"{FILE}/../data/fig2.2/processed/Bachelors2000-23.csv", encoding="utf8")

se_general = pd.read_excel(f"{FILE}/../data/fig2.2/raw/Bachelors1966-2012.xlsx", engine="calamine", skiprows=2, skipfooter=9)

se_general.rename(columns={"Academic year ending": "Year", "All fieldsa": "all fields"}, inplace=True)
se_general = se_general.drop(
  columns=[
    "Unnamed: 2",
    "Mathematics\r\n and computer\r\n sciences",
    "Non-S&E fields",
    "Total",
    "Earth, \r\natmospheric, and ocean sciences",
  ]
)
se_general = pd.melt(se_general, id_vars=["Year"], var_name="Field of Study", value_name="Percent Women")
se_general.dropna(inplace=True)
se_general["Year"] = se_general["Year"].astype(int)

se_compsci = pd.read_excel(f"{FILE}/../data/fig2.2/raw/ComputerSci1966-2012.xlsx", engine="calamine", skiprows=2, skipfooter=9)
se_compsci = se_compsci.iloc[:, :5]
se_compsci.rename(
  columns={"Unnamed: 0": "Year", "All recipients": "Total", "Male": "Men", "Female": "Women"}, inplace=True
)
se_compsci.drop(columns={"Unnamed: 1"}, inplace=True)
se_compsci.dropna(inplace=True)
se_compsci["Percent Men"] = (se_compsci["Men"] / se_compsci["Total"]) * 100
se_compsci["Percent Women"] = (se_compsci["Women"] / se_compsci["Total"]) * 100
se_compsci["Percent Men"] = se_compsci["Percent Men"].round(2)
se_compsci["Percent Women"] = se_compsci["Percent Women"].round(2)
se_compsci["Field of Study"] = "Computer Science"
se_compsci["Year"] = se_general["Year"].astype(int)

se_mathematics = pd.read_excel(
  f"{FILE}/../data/fig2.2/raw/Mathematics1966-2012.xlsx", engine="calamine", skiprows=2, skipfooter=9
)
se_mathematics = se_mathematics.iloc[:, :5]
se_mathematics.rename(
  columns={"Unnamed: 0": "Year", "All recipients": "Total", "Male": "Men", "Female": "Women"}, inplace=True
)
se_mathematics.drop(columns={"Unnamed: 1"}, inplace=True)
se_mathematics.dropna(inplace=True)
se_mathematics["Percent Men"] = (se_mathematics["Men"] / se_mathematics["Total"]) * 100
se_mathematics["Percent Women"] = (se_mathematics["Women"] / se_mathematics["Total"]) * 100
se_mathematics["Percent Men"] = se_mathematics["Percent Men"].round(2)
se_mathematics["Percent Women"] = se_mathematics["Percent Women"].round(2)
se_mathematics["Field of Study"] = "Mathematics"
se_mathematics["Year"] = se_general["Year"].astype(int)

se_combined = pd.concat(
  [
    se_general,
    se_compsci,
    se_mathematics,
  ]
)

se_combined = se_combined[se_combined["Year"] <= 1999]
se_combined["Field of Study"] = se_combined["Field of Study"].str.lower()

# LOL
REMAP_AGAIN = {
  "agriculture": "biological and agricultural sciences",
  "biological sciences": "biological and agricultural sciences",
}

fixed1999_23 = all_years_df.copy()
fixed1999_23["Field of Study"] = fixed1999_23["Field of Study"].replace(REMAP_AGAIN)

finalfinal_combined = pd.concat([se_combined, fixed1999_23])
# print(finalfinal_combined["Field of Study"].unique())

# print(finalfinal_combined.info())
categories_to_show = (
  "all fields",
  "computer science",
  "engineering",
  "biological and agricultural sciences",
  "psychology",
  "mathematics",
  "social sciences",
  "physical sciences",
)
reduced_df = finalfinal_combined.loc[finalfinal_combined["Field of Study"].isin(categories_to_show)]
reduced_df = pd.concat(
  [
    reduced_df,
  ],
  ignore_index=True,
)
left_half = reduced_df[reduced_df["Year"] < 1999]
right_half = reduced_df[reduced_df["Year"] > 1999]
right_half = right_half.groupby(['Year', 'Field of Study']).agg({'Total': 'max', 'Men': 'max', 'Women': 'max'})
right_half['Percent Women'] = (right_half['Women'] / right_half['Total']) * 100
right_half['Percent Men'] = (right_half['Men'] / right_half['Total']) * 100


matplotlib.use("Agg")
sns.lineplot(
  x="Year",
  y="Percent Women",
  hue="Field of Study",
  hue_order=categories_to_show,
  style="Field of Study",
  style_order=categories_to_show,
  data=left_half,
  markers=True,
  dashes=False,
  markersize=8,
)
sns.lineplot(
  x="Year",
  y="Percent Women",
  hue="Field of Study",
  hue_order=categories_to_show,
  style="Field of Study",
  style_order=categories_to_show,
  data=right_half,
  markers=True,
  dashes=False,
  markersize=8,
  legend=False,
)

plt.ylim(0, 100)
plt.xlim(1966, 2023)
plt.xticks(np.arange(1966, 2024, 2), minor=False)
plt.xticks(np.arange(1966, 2024, 1), minor=True)
plt.yticks(np.arange(0, 101, 10))
plt.grid(axis="y", linestyle="--", alpha=0.7)
plt.xticks(rotation=90)
handles, labels = plt.gca().get_legend_handles_labels()

plt.xlabel("Year")
plt.ylabel("Percentage of Women")
plt.title("% of BS degrees awarded to women in the United States, 1966-2023")

plt.legend(handles, labels, title="Field of Study", bbox_to_anchor=(1, 0.75))

# attribution_text = "IPEDs Completions 1999-2023, National Center for Education Statistics (NCES)"

# plt.tight_layout() # Adjust layout to prevent labels from overlapping

# plt.annotate(
#     attribution_text,
#     xy=(0.5, 0.5),  # Position of the annotation (bottom-left corner of axes)
#     xycoords='axes fraction',  # Use axes fractions for positioning
#     xytext=(0.02, 0.02),  # Offset from xy
#     textcoords='axes fraction',
#     bbox=dict(boxstyle="round,pad=0.5", fc="lightgray", ec="gray", lw=1, alpha=0.8), # Box styling
#     fontsize=8,  # Adjust font size
# )
plt.subplots_adjust(bottom=0, left=0, right=1, top=1)
plt.margins(0.10)

plt.savefig(f"{FILE}/../media/fig2.2.png", dpi=300, bbox_inches="tight")
