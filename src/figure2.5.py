import os
from typing import Any
import pandas as pd
import re
import seaborn as sns
import matplotlib
import matplotlib.pyplot as plt
import numpy as np


# https://stackoverflow.com/a/39662359
def is_notebook() -> bool:
  try:
    shell = get_ipython().__class__.__name__
    if shell == "ZMQInteractiveShell":
      return True  # Jupyter notebook or qtconsole
    elif shell == "TerminalInteractiveShell":
      return False  # Terminal running IPython
    else:
      return False  # Other type (?)
  except NameError:
    return False


REMOVE_WHITESPACE = re.compile(r"\s+")
FILE = os.path.dirname(os.path.realpath(__file__)) if not is_notebook() else "."

# fmt: off
se_compsci = pd.read_excel(f"{FILE}/../data/fig2.5/raw/ComputerSci1966-2012.xlsx", engine="calamine", skiprows=2, skipfooter=7)
# fmt: on
se_compsci.rename(
  columns={
    "Unnamed: 0": "year",
    "All recipients": "bachelors_recipients:total",
    "Male": "bachelors_recipients:men",
    "Female": "bachelors_recipients:women",
    "All recipients.1": "masters_recipients:total",
    "Male.1": "masters_recipients:men",
    "Female.1": "masters_recipients:women",
    "All recipients.2": "doctorate_recipients:total",
    "Male.2": "doctorate_recipients:men",
    "Female.2": "doctorate_recipients:women",
  },
  inplace=True,
)
se_compsci.dropna(how="all", axis="columns", inplace=True)
se_compsci.dropna(inplace=True)
se_compsci = se_compsci.astype(int)


def tidy_ipeds(df: pd.DataFrame, year: int) -> pd.DataFrame:
  df_copy = df.copy()

  df_copy.rename(
    columns={
      "Total": "total",
      "Women": "women",
      "Men": "men",
    },
    inplace=True,
  )
  df_copy["Award Level"] = df_copy["Award Level"].apply(lambda x: REMOVE_WHITESPACE.sub(" ", x.strip()).lower())
  df_copy["Award Level"] = df_copy["Award Level"].map(
    {
      "bachelor's degree": "bachelors_recipients",
      "master's degree": "masters_recipients",
      "doctor's degree - research/scholarship": "doctorate_recipients",
      "doctor's degree - professional practicessional": "doctorate_recipients",
      "doctor's degree - other": "doctorate_recipients",
    }
  )
  df_copy = df_copy.groupby(["Award Level"]).agg({"total": "sum", "men": "sum", "women": "sum"})
  df_copy["year"] = year
  df_copy = df_copy.reset_index()
  df_copy = df_copy.pivot(index=["year"], columns=["Award Level"], values=["total", "men", "women"])
  df_copy.columns = df_copy.columns.map(lambda x: ":".join(y for y in x[::-1])).str.strip()
  df_copy["year"] = year

  return df_copy


ipeds_params: dict[str, Any] = {"engine": "calamine", "skiprows": 4, "skipfooter": 4}
# fmt: off
compscigrad2012_13 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2012-13.xlsx", **ipeds_params), 2013)  # type: ignore
compscigrad2013_14 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2013-14.xlsx", **ipeds_params), 2014)  # type: ignore
compscigrad2014_15 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2014-15.xlsx", **ipeds_params), 2015)  # type: ignore
compscigrad2015_16 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2015-16.xlsx", **ipeds_params), 2016)  # type: ignore
compscigrad2016_17 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2016-17.xlsx", **ipeds_params), 2017)  # type: ignore
compscigrad2017_18 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2017-18.xlsx", **ipeds_params), 2018)  # type: ignore
compscigrad2018_19 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2018-19.xlsx", **ipeds_params), 2019)  # type: ignore
compscigrad2019_20 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2019-20.xlsx", **ipeds_params), 2020)  # type: ignore
compscigrad2020_21 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2020-21.xlsx", **ipeds_params), 2021)  # type: ignore
compscigrad2021_22 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2021-22.xlsx", **ipeds_params), 2022)  # type: ignore
compscigrad2022_23 = tidy_ipeds(pd.read_excel(f"{FILE}/../data/fig2.5/raw/CompSciGrad2022-23.xlsx", **ipeds_params), 2023)  # type: ignore
# fmt: on

# se_compsci

aggregated = pd.concat(
  (
    se_compsci,
    compscigrad2012_13,
    compscigrad2013_14,
    compscigrad2014_15,
    compscigrad2015_16,
    compscigrad2016_17,
    compscigrad2017_18,
    compscigrad2018_19,
    compscigrad2019_20,
    compscigrad2020_21,
    compscigrad2021_22,
    compscigrad2022_23,
  )
)
aggregated = aggregated.reset_index(drop=True)
aggregated.to_csv(
  f"{FILE}/../data/fig2.5/processed/Computer Science Degrees Award Level, and Gender Conferred by Year (1966-2023).csv",
  encoding="utf8",
)

id_vars = ["year"]
value_vars = ["bachelors_recipients:total", "masters_recipients:total", "doctorate_recipients:total"]
value_vars_women = ["bachelors_recipients:women", "masters_recipients:women", "doctorate_recipients:women"]
value_vars_men = ["bachelors_recipients:men", "masters_recipients:men", "doctorate_recipients:men"]

melted_df_total = pd.melt(
  aggregated, id_vars=id_vars, value_vars=value_vars, var_name="degree_type", value_name="total"
)
melted_df_women = pd.melt(
  aggregated, id_vars=id_vars, value_vars=value_vars_women, var_name="degree_type", value_name="women"
)
melted_df_men = pd.melt(
  aggregated, id_vars=id_vars, value_vars=value_vars_men, var_name="degree_type", value_name="men"
)

combined_melted = pd.concat(
  (
    melted_df_total,
    melted_df_women,
    melted_df_men,
  )
)

combined_melted["degree_type"] = combined_melted["degree_type"].map(lambda x: x.split(":")[0].split("_")[0])
combined_melted = combined_melted.groupby(["year", "degree_type"], as_index=False).agg(
  {"total": "first", "women": "first", "men": "first"}
)
combined_melted["total"] = combined_melted["total"].astype(int)
combined_melted["women"] = combined_melted["women"].astype(int)
combined_melted["men"] = combined_melted["men"].astype(int)
combined_melted.to_csv(
  f"{FILE}/../data/fig2.5/processed/Computer Science Degrees Conferred by Year, Award Level, and Gender (1966-2023).csv",
  encoding="utf8",
)

left_half = combined_melted[combined_melted["year"] < 1999]
right_half = combined_melted[combined_melted["year"] > 1999]

degree_types = ("bachelors", "masters", "doctorate")

if not is_notebook():
  matplotlib.use("Agg")

sns.lineplot(
  x="year",
  y="total",
  hue="degree_type",
  hue_order=degree_types,
  style="degree_type",
  style_order=degree_types,
  data=left_half,
  markers=True,
  dashes=False,
  markersize=8,
)
axl2 = sns.lineplot(
  x="year",
  y="total",
  hue="degree_type",
  hue_order=degree_types,
  style="degree_type",
  style_order=degree_types,
  data=right_half,
  markers=True,
  dashes=False,
  markersize=8,
  legend=False,
)

plt.xlim(1966, 2023)
plt.ylim(0, 120000)
plt.xticks(np.arange(1966, 2024, 2), minor=False)
plt.xticks(np.arange(1966, 2024, 1), minor=True)
plt.grid(axis="y", linestyle="--", alpha=0.7)
plt.xticks(rotation=90)
handles, labels = plt.gca().get_legend_handles_labels()

plt.xlabel("Year")
plt.ylabel("Degrees Conferred")
plt.title("Computer science degrees conferred in the United States, 1966-2023")

plt.legend(handles, labels, title="")

plt.subplots_adjust(bottom=0, left=0, right=1, top=1)
plt.margins(0.10)

# https://stackoverflow.com/a/62703420
for l in axl2.lines[3:]:  # skip first 3, thats the left half of each line
  y = l.get_ydata()
  if len(y) > 0:
    axl2.annotate(
      f"{y[-1]}", xy=(1.01, y[-1]), xycoords=("axes fraction", "data"), ha="left", va="center", color=l.get_color()
    )

if is_notebook():
  plt.show()

plt.savefig(f"{FILE}/../media/fig2.5.png", dpi=300, bbox_inches="tight")
