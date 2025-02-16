import os
from typing import Any, cast
import pandas as pd
import re
import seaborn as sns
import matplotlib
import matplotlib.pyplot as plt
import numpy as np


# https://stackoverflow.com/a/39662359
def is_notebook() -> bool:
  try:
    shell = get_ipython().__class__.__name__  # type: ignore
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

compsci_all1966_2023 = pd.read_csv(  # type: ignore
  f"{FILE}/../data/fig2.5/processed/Computer Science Degrees Conferred by Year, Award Level, and Gender (1966-2023).csv",
  index_col=0,
)

compsci_all1966_2023 = compsci_all1966_2023.rename(
  columns={"total": "compsci:total", "women": "compsci:women", "men": "compsci:men"}
)

# fmt: off
se_general = pd.read_excel(f"{FILE}/../data/fig13.2/raw/AwardLevelandSex1966-2012.xlsx", engine="calamine", skiprows=2, skipfooter=5) # type: ignore
# fmt: on
se_general.rename(
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
se_general.columns
se_general.drop(columns=["% female", "% female.1", "% female.2"], inplace=True)
se_general.dropna(how="all", axis="columns", inplace=True)  # type: ignore
se_general.dropna(inplace=True, how="all")  # type: ignore
se_general["year"] = se_general["year"].astype(int)
se_general["bachelors_recipients:total"] = (
  se_general["bachelors_recipients:women"] + se_general["bachelors_recipients:men"]
)
se_general["masters_recipients:total"] = se_general["masters_recipients:women"] + se_general["masters_recipients:men"]
se_general["doctorate_recipients:total"] = (
  se_general["doctorate_recipients:women"] + se_general["doctorate_recipients:men"]
)

# Doctoral has data for 1999, but others are missing
se_general["doctorate_recipients:total"] = se_general["doctorate_recipients:total"].astype(int)
se_general["doctorate_recipients:men"] = se_general["doctorate_recipients:men"].astype(int)
se_general["doctorate_recipients:women"] = se_general["doctorate_recipients:women"].astype(int)

id_vars = ["year"]
value_vars = ["bachelors_recipients:total", "masters_recipients:total", "doctorate_recipients:total"]
value_vars_women = ["bachelors_recipients:women", "masters_recipients:women", "doctorate_recipients:women"]
value_vars_men = ["bachelors_recipients:men", "masters_recipients:men", "doctorate_recipients:men"]

melted_df_total = pd.melt(  # type: ignore
  se_general, id_vars=id_vars, value_vars=value_vars, var_name="degree_type", value_name="total"
)
melted_df_women = pd.melt(  # type: ignore
  se_general, id_vars=id_vars, value_vars=value_vars_women, var_name="degree_type", value_name="women"
)
melted_df_men = pd.melt(  # type: ignore
  se_general, id_vars=id_vars, value_vars=value_vars_men, var_name="degree_type", value_name="men"
)
combined_melted = pd.concat(
  (
    melted_df_total,
    melted_df_women,
    melted_df_men,
  )
)

combined_melted["degree_type"] = combined_melted["degree_type"].map(lambda x: x.split(":")[0].split("_")[0])  # type: ignore
combined_melted = combined_melted.groupby(["year", "degree_type"], as_index=False).agg(  # type: ignore
  {"total": "first", "women": "first", "men": "first"}
)
combined_melted.dropna(inplace=True, how="any")  # type: ignore
combined_melted["total"] = combined_melted["total"].astype(int)
combined_melted["women"] = combined_melted["women"].astype(int)
combined_melted["men"] = combined_melted["men"].astype(int)
combined_melted.rename(columns={"total": "all:total", "women": "all:women", "men": "all:men"}, inplace=True)

general_params: dict[str, Any] = {"engine": "calamine", "skiprows": 4, "skipfooter": 5}
# fmt: off
bachelors2012_23 = cast(pd.DataFrame, pd.read_excel(f"{FILE}/../data/fig13.2/raw/BachelorsCompletions2012-23.xlsx", **general_params) ) # type: ignore
masters2012_23   = cast(pd.DataFrame, pd.read_excel(f"{FILE}/../data/fig13.2/raw/MastersCompletions2012-23.xlsx", **general_params)   ) # type: ignore
doctors2012_23a  = cast(pd.DataFrame, pd.read_excel(f"{FILE}/../data/fig13.2/raw/DoctorsCompletionsOth2012-23.xlsx", **general_params)) # type: ignore
doctors2012_23b  = cast(pd.DataFrame, pd.read_excel(f"{FILE}/../data/fig13.2/raw/DoctorsCompletionsPro2012-23.xlsx", **general_params)) # type: ignore
doctors2012_23c  = cast(pd.DataFrame, pd.read_excel(f"{FILE}/../data/fig13.2/raw/DoctorsCompletionsRes2012-23.xlsx", **general_params)) # type: ignore
# fmt: on


def tidy_year_data(d: pd.DataFrame, degree_type: str) -> pd.DataFrame:
  df = d.copy()
  df = df.iloc[:, 1:].T
  df.columns = df.iloc[0]
  df = df[1:].reset_index()
  df.rename(columns={"index": "year", "Total": "all:total", "Men": "all:men", "Women": "all:women"}, inplace=True)
  df["year"] = df["year"].map(lambda x: f"{x[:2]}{x[-2:]}")  # type: ignore
  df = df.astype(int)  # type: ignore
  df["degree_type"] = degree_type
  return df


bachelors2012_23 = tidy_year_data(bachelors2012_23, "bachelors")
masters2012_23 = tidy_year_data(masters2012_23, "masters")
doctors2012_23 = (
  pd.concat(
    (
      tidy_year_data(doctors2012_23a, "doctorate"),
      tidy_year_data(doctors2012_23b, "doctorate"),
      tidy_year_data(doctors2012_23c, "doctorate"),
    )
  )
  .groupby(["year", "degree_type"], as_index=False)  # type: ignore
  .agg({"all:total": "sum", "all:women": "sum", "all:men": "sum"})
)

aggregate = (
  pd.concat((compsci_all1966_2023, combined_melted, bachelors2012_23, masters2012_23, doctors2012_23))
  .groupby(["year", "degree_type"], as_index=False)  # type: ignore
  .agg(
    {
      "all:total": "first",
      "all:women": "first",
      "all:men": "first",
      "compsci:total": "first",
      "compsci:women": "first",
      "compsci:men": "first",
    }
  )
)
aggregate = aggregate.astype({col: int for col in aggregate.columns[2:]})  # type: ignore
aggregate["compsci_percent:total"] = ((aggregate["compsci:total"] / aggregate["all:total"]) * 100).round(2)  # type: ignore
aggregate["compsci_percent:women"] = ((aggregate["compsci:women"] / aggregate["all:women"]) * 100).round(2)  # type: ignore
aggregate["compsci_percent:men"] = ((aggregate["compsci:men"] / aggregate["all:men"]) * 100).round(2)  # type: ignore

aggregate.to_csv(
  f"{FILE}/../data/fig13.2/processed/Proportion of Undergraduates earning Computer Science Degrees by Gender, 1966-2023.csv",
  encoding="utf8",
)

left_half = aggregate[aggregate["year"] < 1999]
right_half = aggregate[aggregate["year"] >= 1999]

degree_types = ("bachelors", "masters", "doctorate")

if not is_notebook():
  matplotlib.use("Agg")

sns.lineplot(
  x="year",
  y="compsci_percent:women",
  hue="degree_type",
  hue_order=("bachelors",),
  style="degree_type",
  style_order=("bachelors",),
  data=left_half,
  markers=True,
  markersize=8,
)
sns.lineplot(
  x="year",
  y="compsci_percent:men",
  hue="degree_type",
  hue_order=(
    None,
    "bachelors",
  ),  # type: ignore
  style="degree_type",
  style_order=(
    None,
    "bachelors",
  ),  # type: ignore
  data=left_half,
  markers=True,
  markersize=8,
)
axl2 = sns.lineplot(
  x="year",
  y="compsci_percent:women",
  hue="degree_type",
  hue_order=("bachelors",),
  style="degree_type",
  style_order=("bachelors",),
  data=right_half,
  markers=True,
  markersize=8,
  legend=False,
)
sns.lineplot(
  x="year",
  y="compsci_percent:men",
  hue="degree_type",
  hue_order=(
    None,
    "bachelors",
  ),  # type: ignore
  style="degree_type",
  style_order=(
    None,
    "bachelors",
  ),  # type: ignore
  data=right_half,
  markers=True,
  markersize=8,
  legend=False,
)

plt.xlim(1966, 2023)  # type: ignore
plt.ylim(0, 12)  # type: ignore
plt.xticks(np.arange(1966, 2024, 2), minor=False)  # type: ignore
plt.xticks(np.arange(1966, 2024, 1), minor=True)  # type: ignore
plt.grid(axis="y", linestyle="--", alpha=0.7)  # type: ignore
plt.xticks(rotation=90)  # type: ignore
handles, labels = plt.gca().get_legend_handles_labels()

plt.xlabel("Year")  # type: ignore
plt.ylabel("% of Gender Graduating with CS BS Degree")  # type: ignore
plt.title("Proportion of CS Graduates by Gender in the United States, 1966-2023")  # type: ignore

plt.legend(  # type: ignore
  handles=handles,
  labels=[
    "% women in CS as a % of all female bachelors",
    "% men in CS as a % of all male bachelors",
  ],
  title="",
  loc=(0.01, 0.89),
)

plt.subplots_adjust(bottom=0, left=0, right=1, top=1)
plt.margins(0.10)

# https://stackoverflow.com/a/62703420
for l in axl2.lines[3:]:  # skip first 3, thats the left half of each line  # noqa: E741
  y = l.get_ydata()
  if len(y) > 0:  # type: ignore
    axl2.annotate(  # type: ignore
      f"{y[-1]}%",  # type: ignore
      xy=(1.01, y[-1]),  # type: ignore
      xycoords=("axes fraction", "data"),
      ha="left",
      va="center",
      color=l.get_color(),  # type: ignore
    )

colors: list[Any] = []
for l in axl2.lines[3:]:  # noqa: E741
  colors.append(l.get_color())

# print(colors)

offset_x = 0.5
offset_y = 0.3
axl2.annotate("5.52%", xy=(1986 + offset_x, 5.52 + offset_y), textcoords="data", ha="center", color=colors[0])  # type: ignore
axl2.annotate("2.97%", xy=(1986 + offset_x, 2.97 + offset_y), textcoords="data", ha="center", color=colors[1])  # type: ignore

axl2.annotate("7.46%", xy=(2004 + offset_x, 7.46 + offset_y), textcoords="data", ha="center", color=colors[0])  # type: ignore
axl2.annotate("1.99%", xy=(2003 + offset_x, 1.99 + offset_y), textcoords="data", ha="center", color=colors[1])  # type: ignore

plt.axvline(1986, 0, 2.97 / 12, dashes=[1, 2], color=colors[1])  # type: ignore
plt.axvline(1986, 3.75 / 12, 5.52 / 12, dashes=[1, 4], color=colors[0])  # type: ignore
plt.axvline(2003, 0, 1.99 / 12, dashes=[1, 2], color=colors[1])  # type: ignore
plt.axvline(2004, 2.75 / 12, 7.46 / 12, dashes=[1, 4], color=colors[0])  # type: ignore
plt.axvline(2004, 0, (1.85 - 0.1) / 12, dashes=[1, 4], color=colors[0])  # type: ignore

if is_notebook():
  plt.show()  # type: ignore

plt.savefig(f"{FILE}/../media/fig13.2.png", dpi=300, bbox_inches="tight")  # type: ignore

aggregate  # type: ignore
