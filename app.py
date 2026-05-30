import streamlit as st
import numpy as np
import pandas as pd
import io
import xlsxwriter
from scipy.optimize import linprog

# -------------------------------
# 1. SESSION STATE INITIALIZATION
# -------------------------------
if 'species_idx' not in st.session_state:
    st.session_state.species_idx = 0
if 'method_idx' not in st.session_state:
    st.session_state.method_idx = 0
if 'stage_idx' not in st.session_state:
    st.session_state.stage_idx = 0
if 'manual_cp' not in st.session_state:
    st.session_state.manual_cp = 30.0
if 'feed_qty' not in st.session_state:
    st.session_state.feed_qty = 100.0

# -------------------------------
# 2. NUTRITIONAL DATABASE (AS‑FED)
# -------------------------------
INGREDIENTS = {
    "Soybean Meal": {
        "CP": 44, "Fat": 1.5, "Fiber": 6, "Moisture": 10, "Ash": 6,
        "Lys": 2.7, "Iso": 2.1, "Max": 0.35, "Min": 0.0, "Source": "Plant"
    },
    "Peanut Meal/Ambaz": {
        "CP": 45, "Fat": 6, "Fiber": 7, "Moisture": 7, "Ash": 5.5,
        "Lys": 1.6, "Iso": 1.5, "Max": 0.35, "Min": 0.0, "Source": "Plant"
    },
    "Sunflower Meal/Ambaz": {
        "CP": 36, "Fat": 2.5, "Fiber": 14, "Moisture": 8, "Ash": 6.5,
        "Lys": 1.2, "Iso": 1.4, "Max": 0.20, "Min": 0.0, "Source": "Plant"
    },
    "Cottonseed Meal/Ambaz": {
        "CP": 38, "Fat": 2, "Fiber": 11, "Moisture": 9, "Ash": 6.5,
        "Lys": 1.4, "Iso": 1.3, "Max": 0.15, "Min": 0.0, "Source": "Plant"
    },
    "Local Fish Meal": {
        "CP": 55, "Fat": 9, "Fiber": 1, "Moisture": 8, "Ash": 18,
        "Lys": 4.1, "Iso": 2.6, "Max": 0.25, "Min": 0.0, "Source": "Animal",
        "Fry_Min": 0.05
    },
    "Poultry By-product Meal": {
        "CP": 58, "Fat": 12, "Fiber": 1.5, "Moisture": 6, "Ash": 12,
        "Lys": 3.2, "Iso": 2.3, "Max": 0.25, "Min": 0.0, "Source": "Animal"
    },
    "Yellow Corn": {
        "CP": 9, "Fat": 3.8, "Fiber": 2.5, "Moisture": 12, "Ash": 1.5,
        "Lys": 0.24, "Iso": 0.31, "Max": 0.50, "Min": 0.0, "Source": "Plant"
    },
    "Sorghum Feterita": {
        "CP": 10, "Fat": 3.2, "Fiber": 2.7, "Moisture": 11, "Ash": 1.8,
        "Lys": 0.22, "Iso": 0.38, "Max": 0.50, "Min": 0.0, "Source": "Plant"
    },
    "Wheat Bran/Rada": {
        "CP": 15, "Fat": 4, "Fiber": 10, "Moisture": 11, "Ash": 5,
        "Lys": 0.58, "Iso": 0.51, "Max": 0.25, "Min": 0.0, "Source": "Plant"
    },
    "Rice Bran/Rajee' Al-Koun": {
        "CP": 12, "Fat": 12, "Fiber": 11.5, "Moisture": 10, "Ash": 9,
        "Lys": 0.5, "Iso": 0.45, "Max": 0.20, "Min": 0.0, "Source": "Plant"
    },
    "Supplemental Lipids/Oils": {
        "CP": 0, "Fat": 100, "Fiber": 0, "Moisture": 0, "Ash": 0,
        "Lys": 0, "Iso": 0, "Max": 0.06, "Min": 0.0, "Source": "Plant"
    }
}
INGREDIENT_KEYS = list(INGREDIENTS.keys())

# Initialize selected_ingredients after INGREDIENT_KEYS is defined
if 'selected_ingredients' not in st.session_state:
    st.session_state.selected_ingredients = {ing: False for ing in INGREDIENT_KEYS}

# -------------------------------
# 3. SPECIES‑SPECIFIC PROFILES
# -------------------------------
SPECIES_PROFILES = {
    "Nile Tilapia": {
        "stages": ["Fry (<10g)", "Fingerling (10-30g)", "Grower (30-150g)", "Finisher (>150g)"],
        "profiles": {
            "Fry (<10g)": {
                "target_cp": 40, "max_fat": 10, "max_fiber": 5,
                "pellet_size": "1.0 mm crumble",
                "lys_pct_of_cp": 5.5, "iso_pct_of_cp": 3.5
            },
            "Fingerling (10-30g)": {
                "target_cp": 35, "max_fat": 10, "max_fiber": 6,
                "pellet_size": "2.0 mm pellet",
                "lys_pct_of_cp": 5.5, "iso_pct_of_cp": 3.5
            },
            "Grower (30-150g)": {
                "target_cp": 30, "max_fat": 8, "max_fiber": 7,
                "pellet_size": "3.0 mm pellet",
                "lys_pct_of_cp": 5.5, "iso_pct_of_cp": 3.5
            },
            "Finisher (>150g)": {
                "target_cp": 28, "max_fat": 8, "max_fiber": 8,
                "pellet_size": "4.0 mm pellet",
                "lys_pct_of_cp": 5.5, "iso_pct_of_cp": 3.5
            }
        }
    },
    "African Catfish": {
        "stages": ["Fry (<5g)", "Fingerling (5-20g)", "Grower (20-200g)", "Finisher (>200g)"],
        "profiles": {
            "Fry (<5g)": {
                "target_cp": 45, "max_fat": 12, "max_fiber": 4,
                "pellet_size": "1.0 mm crumble",
                "lys_pct_of_cp": 5.1, "iso_pct_of_cp": 4.0
            },
            "Fingerling (5-20g)": {
                "target_cp": 40, "max_fat": 12, "max_fiber": 5,
                "pellet_size": "2.0 mm pellet",
                "lys_pct_of_cp": 5.1, "iso_pct_of_cp": 4.0
            },
            "Grower (20-200g)": {
                "target_cp": 35, "max_fat": 14, "max_fiber": 6,
                "pellet_size": "3.0 mm pellet",
                "lys_pct_of_cp": 5.1, "iso_pct_of_cp": 4.0
            },
            "Finisher (>200g)": {
                "target_cp": 30, "max_fat": 16, "max_fiber": 7,
                "pellet_size": "4.0 mm pellet",
                "lys_pct_of_cp": 5.1, "iso_pct_of_cp": 4.0
            }
        }
    }
}

# -------------------------------
# 4. LINEAR PROGRAMMING SOLVER
# -------------------------------
def solve_formulation(selected_ingredients, target_cp, max_fat, max_fiber, stage_is_fry=False):
    n = len(selected_ingredients)
    if n == 0:
        return None

    A_eq = np.zeros((2, n))
    b_eq = np.zeros(2)
    A_eq[0, :] = 1.0
    b_eq[0] = 0.98
    cp_vals = [INGREDIENTS[ing]["CP"] for ing in selected_ingredients]
    A_eq[1, :] = cp_vals
    b_eq[1] = target_cp

    A_ub = np.zeros((2, n))
    b_ub = np.zeros(2)
    fat_vals = [INGREDIENTS[ing]["Fat"] for ing in selected_ingredients]
    fiber_vals = [INGREDIENTS[ing]["Fiber"] for ing in selected_ingredients]
    A_ub[0, :] = fat_vals
    b_ub[0] = max_fat
    A_ub[1, :] = fiber_vals
    b_ub[1] = max_fiber

    bounds = []
    for ing in selected_ingredients:
        lb = 0.0
        ub = INGREDIENTS[ing]["Max"]
        if ing == "Local Fish Meal" and stage_is_fry:
            lb = INGREDIENTS[ing].get("Fry_Min", 0.0)
        bounds.append((lb, ub))

    c = np.zeros(n)
    res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq,
                  bounds=bounds, method='highs')

    if res.success:
        return res.x
    else:
        return None

# -------------------------------
# 5. AMINO ACID DEFICIT CALCULATION
# -------------------------------
def compute_amino_supplements(selected_ingredients, fractions, target_cp, species):
    if species == "Nile Tilapia":
        lys_pct_of_cp = 5.5
        iso_pct_of_cp = 3.5
    else:
        lys_pct_of_cp = 5.1
        iso_pct_of_cp = 4.0

    target_lys_pct = target_cp * lys_pct_of_cp / 100.0
    target_iso_pct = target_cp * iso_pct_of_cp / 100.0

    lys_g_kg = sum([INGREDIENTS[ing]["Lys"] * frac * 10.0 for ing, frac in zip(selected_ingredients, fractions)])
    iso_g_kg = sum([INGREDIENTS[ing]["Iso"] * frac * 10.0 for ing, frac in zip(selected_ingredients, fractions)])

    lys_req_g_kg = target_lys_pct * 10.0
    iso_req_g_kg = target_iso_pct * 10.0

    lys_deficit = max(0.0, lys_req_g_kg - lys_g_kg)
    iso_deficit = max(0.0, iso_req_g_kg - iso_g_kg)

    lys_hcl_kg_per_kg = lys_deficit / 788.0   # g Lys deficit / 788 g Lys per kg L-Lys HCl
    iso_kg_per_kg = iso_deficit / 1000.0

    return {
        "lys_deficit_g_kg": lys_deficit,
        "iso_deficit_g_kg": iso_deficit,
        "lys_hcl_kg_per_kg": lys_hcl_kg_per_kg,
        "iso_kg_per_kg": iso_kg_per_kg
    }

# -------------------------------
# 6. OUTPUT GENERATION (3 decimals always)
# -------------------------------
def generate_batch_sheet(selected_ingredients, fractions, total_kg, supplement_info):
    rows = []
    for ing, frac in zip(selected_ingredients, fractions):
        kg = frac * total_kg
        rows.append({
            "Ingredient": ing,
            "Inclusion (%)": round(frac * 100, 3),
            "Weight (kg)": round(kg, 3)
        })
    # Fixed premixes
    rows.append({
        "Ingredient": "Vitamin Premix (fixed 1%)",
        "Inclusion (%)": 1.000,
        "Weight (kg)": round(0.01 * total_kg, 3)
    })
    rows.append({
        "Ingredient": "Mineral Premix (fixed 1%)",
        "Inclusion (%)": 1.000,
        "Weight (kg)": round(0.01 * total_kg, 3)
    })
    # Amino acid supplements
    if supplement_info["lys_hcl_kg_per_kg"] > 1e-6:
        kg_lys = supplement_info["lys_hcl_kg_per_kg"] * total_kg
        rows.append({
            "Ingredient": "L-Lysine HCl (78.8% Lys)",
            "Inclusion (%)": round(supplement_info["lys_hcl_kg_per_kg"] * 100, 3),
            "Weight (kg)": round(kg_lys, 3)
        })
    if supplement_info["iso_kg_per_kg"] > 1e-6:
        kg_iso = supplement_info["iso_kg_per_kg"] * total_kg
        rows.append({
            "Ingredient": "L-Isoleucine (pure)",
            "Inclusion (%)": round(supplement_info["iso_kg_per_kg"] * 100, 3),
            "Weight (kg)": round(kg_iso, 3)
        })
    df = pd.DataFrame(rows)
    # Add total row
    total_weight = df["Weight (kg)"].sum()
    total_row = pd.DataFrame([["", "", round(total_weight, 3)]], columns=df.columns)
    df = pd.concat([df, total_row], ignore_index=True)
    df.iloc[-1, 0] = "TOTAL"
    return df

def proximate_analysis(selected_ingredients, fractions):
    cp = sum(INGREDIENTS[ing]["CP"] * frac for ing, frac in zip(selected_ingredients, fractions))
    fat = sum(INGREDIENTS[ing]["Fat"] * frac for ing, frac in zip(selected_ingredients, fractions))
    fiber = sum(INGREDIENTS[ing]["Fiber"] * frac for ing, frac in zip(selected_ingredients, fractions))
    moist = sum(INGREDIENTS[ing]["Moisture"] * frac for ing, frac in zip(selected_ingredients, fractions))
    ash = sum(INGREDIENTS[ing]["Ash"] * frac for ing, frac in zip(selected_ingredients, fractions))
    return cp, fat, fiber, moist, ash

def plant_animal_ratio(selected_ingredients, fractions):
    plant_prot = sum(INGREDIENTS[ing]["CP"] * frac for ing, frac in zip(selected_ingredients, fractions)
                     if INGREDIENTS[ing]["Source"] == "Plant")
    animal_prot = sum(INGREDIENTS[ing]["CP"] * frac for ing, frac in zip(selected_ingredients, fractions)
                      if INGREDIENTS[ing]["Source"] == "Animal")
    total_prot = plant_prot + animal_prot
    if total_prot > 0:
        plant_ratio = plant_prot / total_prot * 100
        animal_ratio = animal_prot / total_prot * 100
    else:
        plant_ratio = animal_ratio = 0.0
    return plant_ratio, animal_ratio

def starch_fraction(selected_ingredients, fractions):
    starch_ings = {"Yellow Corn", "Sorghum Feterita", "Wheat Bran/Rada", "Rice Bran/Rajee' Al-Koun"}
    total = 0.0
    for ing, frac in zip(selected_ingredients, fractions):
        if ing in starch_ings:
            total += frac
    return total

# -------------------------------
# 7. EXCEL EXPORT (3 decimals)
# -------------------------------
def create_excel(df_batch, meta_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_batch.to_excel(writer, sheet_name='Batch Formulation', index=False)
        meta_df = pd.DataFrame(meta_data.items(), columns=['Parameter', 'Value'])
        meta_df.to_excel(writer, sheet_name='Meta Summary', index=False)

        workbook = writer.book
        # Format for 3 decimals
        fmt_3dec = workbook.add_format({'num_format': '0.000'})
        for sheet_name in writer.sheets:
            sheet = writer.sheets[sheet_name]
            sheet.set_column(0, 0, 30)  # First column wider
            if sheet_name == 'Batch Formulation':
                # Set numeric columns format
                sheet.set_column(1, 1, 15, fmt_3dec)  # Inclusion
                sheet.set_column(2, 2, 15, fmt_3dec)  # Weight
            elif sheet_name == 'Meta Summary':
                sheet.set_column(1, 1, 15, fmt_3dec)
    return output.getvalue()

# -------------------------------
# 8. STREAMLIT UI
# -------------------------------
st.set_page_config(page_title="Fish Feed Formulator", layout="wide")
st.title("🐟 Freshwater Fish Feed Formulation Tool")
st.markdown("#### Nile Tilapia (*Oreochromis niloticus*) & African Catfish (*Clarias lazera*)")
st.markdown("---")

# Sidebar reset
with st.sidebar:
    if st.button("🔄 Reset All Parameters"):
        for key in ['species_idx', 'method_idx', 'stage_idx', 'manual_cp', 'feed_qty', 'selected_ingredients']:
            if key in st.session_state:
                if key == 'selected_ingredients':
                    st.session_state.selected_ingredients = {ing: False for ing in INGREDIENT_KEYS}
                elif key == 'manual_cp':
                    st.session_state.manual_cp = 30.0
                elif key == 'feed_qty':
                    st.session_state.feed_qty = 100.0
                else:
                    st.session_state[key] = 0
        st.success("All settings reset to defaults.")

# Species & Method selection
col1, col2 = st.columns(2)
with col1:
    species_list = list(SPECIES_PROFILES.keys())
    species = st.selectbox("1️⃣ Fish Species", species_list,
                           index=st.session_state.species_idx,
                           key='species_sel')
    st.session_state.species_idx = species_list.index(species)
with col2:
    method = st.radio("2️⃣ Protein Targeting Method",
                      ["Automated by Weight/Growth Stage", "Direct Manual Entry"],
                      index=st.session_state.method_idx,
                      key='method_sel')
    st.session_state.method_idx = 0 if method.startswith("Automated") else 1

# Target CP and limits
if method.startswith("Automated"):
    profile = SPECIES_PROFILES[species]
    stage = st.selectbox("Growth Stage", profile["stages"],
                         index=st.session_state.stage_idx,
                         key='stage_sel')
    st.session_state.stage_idx = profile["stages"].index(stage)
    stage_data = profile["profiles"][stage]
    target_cp = stage_data["target_cp"]
    max_fat = stage_data["max_fat"]
    max_fiber = stage_data["max_fiber"]
    pellet_advice = stage_data["pellet_size"]
    st.info(f"**Pellet size advice:** {pellet_advice}  \n"
            f"**Target CP:** {target_cp}%  |  Max Fat: {max_fat}%  |  Max Fiber: {max_fiber}%")
else:
    target_cp = st.number_input("Target Crude Protein (%)", min_value=10.0, max_value=60.0,
                                value=st.session_state.manual_cp, step=0.5,
                                key='manual_cp_input')
    st.session_state.manual_cp = target_cp
    if species == "Nile Tilapia":
        max_fat, max_fiber = 12.0, 7.0
    else:
        max_fat, max_fiber = 16.0, 6.0

feed_kg = st.number_input("Total Feed Quantity (kg)", min_value=1.0, value=st.session_state.feed_qty, step=1.0,
                          key='feed_qty_input')
st.session_state.feed_qty = feed_kg

# Ingredient selection
st.subheader("📦 Available Ingredients (Check those in warehouse)")
cols = st.columns(3)
for i, ing in enumerate(INGREDIENT_KEYS):
    with cols[i % 3]:
        st.session_state.selected_ingredients[ing] = st.checkbox(
            f"{ing} (CP:{INGREDIENTS[ing]['CP']}%, Fat:{INGREDIENTS[ing]['Fat']}%, Fiber:{INGREDIENTS[ing]['Fiber']}%)",
            value=st.session_state.selected_ingredients[ing],
            key=f"cb_{ing}"
        )

selected_ings = [ing for ing, checked in st.session_state.selected_ingredients.items() if checked]

st.markdown("---")

if st.button("⚙️ Formulate Feed", type="primary"):
    if not selected_ings:
        st.error("Please select at least one ingredient.")
    else:
        is_fry = method.startswith("Automated") and "Fry" in stage
        fractions = solve_formulation(selected_ings, target_cp, max_fat, max_fiber, stage_is_fry=is_fry)

        if fractions is None:
            st.error("❌ **Infeasible solution**: The current ingredient set cannot achieve "
                     "the target protein with the given fat/fiber limits. "
                     "Consider adding concentrated protein sources (e.g., Local Fish Meal, "
                     "Poultry By‑product Meal) or relaxing constraints via manual entry.")
            st.stop()

        suppl = compute_amino_supplements(selected_ings, fractions, target_cp, species)
        cp_act, fat_act, fib_act, moist_act, ash_act = proximate_analysis(selected_ings, fractions)
        plant_r, animal_r = plant_animal_ratio(selected_ings, fractions)
        starch_frac = starch_fraction(selected_ings, fractions)
        binder_warning = starch_frac < 0.25
        cot_frac = fractions[selected_ings.index("Cottonseed Meal/Ambaz")] if "Cottonseed Meal/Ambaz" in selected_ings else 0.0

        # Display dashboard with 3 decimals
        st.success("✅ Optimal formulation found!")
        st.subheader("📊 Proximate Analysis Dashboard")
        c1, c2, c3 = st.columns(3)
        c1.metric("Crude Protein (%)", f"{cp_act:.3f}", delta=f"Target {target_cp}%")
        c1.metric("Crude Fat (%)", f"{fat_act:.3f}")
        c2.metric("Crude Fiber (%)", f"{fib_act:.3f}")
        c2.metric("Total Moisture (%)", f"{moist_act:.3f}")
        c3.metric("Ash (%)", f"{ash_act:.3f}")
        c3.metric("Plant Protein Ratio", f"{plant_r:.1f}%", delta=f"Animal {animal_r:.1f}%")

        if moist_act > 11.5:
            st.warning("⚠️ **High moisture (>11.5%)** – aflatoxin risk in hot climates. Ensure immediate drying and proper storage.")
        if fib_act > 7.0:
            st.warning("⚠️ **High fiber (>7%)** may reduce growth performance and affect biofloc water quality.")
        if cot_frac > 0.12:
            st.warning("⚠️ **Cottonseed meal >12%** – risk of gossypol toxicity. Consider limiting or using iron‑treated meal.")
        if binder_warning:
            st.warning("💧 **Water stability alert:** Total starch sources are below 25%. "
                       "Add 1‑2% commercial binder or pre‑gelatinized starch to prevent pellet disintegration.")

        # Batch sheet
        df_batch = generate_batch_sheet(selected_ings, fractions, feed_kg, suppl)
        st.subheader("📋 Batch Formulation Sheet")
        st.dataframe(df_batch.style.format({"Inclusion (%)": "{:.3f}", "Weight (kg)": "{:.3f}"}))

        # Amino acid supplement summary
        if suppl["lys_hcl_kg_per_kg"] > 1e-6 or suppl["iso_kg_per_kg"] > 1e-6:
            st.markdown("### 🔬 Amino Acid Fortification")
            if suppl["lys_deficit_g_kg"] > 0:
                st.write(f"Lysine deficit: {suppl['lys_deficit_g_kg']:.3f} g/kg feed → "
                         f"add **{suppl['lys_hcl_kg_per_kg']*feed_kg:.3f} kg** L‑Lysine HCl per batch.")
            if suppl["iso_deficit_g_kg"] > 0:
                st.write(f"Isoleucine deficit: {suppl['iso_deficit_g_kg']:.3f} g/kg feed → "
                         f"add **{suppl['iso_kg_per_kg']*feed_kg:.3f} kg** L‑Isoleucine per batch.")

        # Meta summary and Excel download
        meta = {
            "Target CP (%)": target_cp,
            "Actual CP (%)": round(cp_act, 3),
            "Actual Fat (%)": round(fat_act, 3),
            "Actual Fiber (%)": round(fib_act, 3),
            "Moisture (%)": round(moist_act, 3),
            "Ash (%)": round(ash_act, 3),
            "Plant Protein Ratio (%)": round(plant_r, 1),
            "Animal Protein Ratio (%)": round(animal_r, 1),
            "Starch Fraction (%)": round(starch_frac*100, 3),
            "Cottonseed Meal Inclusion (%)": round(cot_frac*100, 3)
        }
        excel_data = create_excel(df_batch, meta)

        st.download_button(
            label="📥 Download Batch Sheet (.xlsx)",
            data=excel_data,
            file_name="fish_feed_formulation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")
st.caption("Aquaculture Nutrition Expert System • Cloud‑Ready • LP Solver with scipy.optimize")
