""" ECOHUB - A Personal Carbon Footprint Tracker Application using Tkinter """

"""
Code Section Label Meanings:
EXPENSEWISE: Code sections derived from a Finance Tracking App (LBYCPA1 project of Zander Ivo Holoyohoy).
             Includes most Utility Functions, Account Management (AccountsPage, Add Profile,
             SimpleEntryDialog, Switch/Reset/Delete User, Exit App), Theme Switching, and Sidebar.
IVO-ONLY:    Core backend logic, calculations, and data persistence handled solely by Zander Ivo Holoyohoy.
             Includes Base Categories, Data Store structure, File Paths, Data Loading/Saving,
             non-UI parts of AddCarbonFootprintActivityDialog (calculation, validation, submission helpers),
             and the main application launch logic.
IVO+GPT:     GUI elements, object-oriented programming architecture, and visual components where Zander Ivo Holoyohoy received assistance,
             primarily focusing on Tkinter implementation beyond the scope of the PROLOGI topics.
             Includes BasePage structure, specific category page layouts (Residential, Travel, etc.),
             CarbonDashboardPage layout, UI widget creation within AddCarbonFootprintActivityDialog,
             and other UI-heavy functions/classes. Specifically used Gemini 2.5 Pro Experimental for GUI & OOP Code Generation
             + Claude 3.7 Sonnet for Very Minor Debugging, Error Detection, and Code Optimizations.
GUTIERREZ+KATSUYA: Specific numerical data, emission factors, and conversion logic derived
                   from research by Gutierrez & Katsuya. Primarily the DEFAULT_EMISSION_FACTORS
                   and the CO2e to Trees/Cars conversion logic.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any
import datetime
import random
import csv
import os
import json
import logging

# --- Logging Setup ---
# Basic configuration logs INFO level messages to console/file
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
)

# IVO+GPT
# --- Configuration ---
# Theme Definitions (Using constants for keys improves maintainability)
BG = "background"
FG = "foreground"
SIDEBAR = "sidebar"
CARD = "card"
ACCENT = "accent"
ACCENT_DARKER = "accent_darker"
RED = "red"
BLUE = "blue" # Unused? Keep for potential future use.
YELLOW = "yellow" # Unused? Keep for potential future use.
DISABLED = "disabled"
ACC_ICON_BG = "account_icon_bg"
DLG_BG = "dialog_bg"
DLG_FG = "dialog_fg"
DLG_CARD = "dialog_card"
TV_HEAD_BG = "treeview_heading_bg"
TV_ODD = "treeview_odd_row"
TV_EVEN = "treeview_even_row"
SB_TROUGH = "scrollbar_trough"
SB_BG = "scrollbar_bg"
BTN_FG = "button_fg"
CB_LIST_BG = "combobox_list_bg"
CB_LIST_FG = "combobox_list_fg"
CB_LIST_SEL_BG = "combobox_list_select_bg"
CB_LIST_SEL_FG = "combobox_list_select_fg"
LBL_DESC_FG = "label_desc_fg"

# IVO+GPT (IVO-HEXCODE SELECTION, GPT-HEXCODE REPETITIONS)
THEME_ECO_DARK = {
    BG: "#242938", FG: "#c5cae9", SIDEBAR: "#30384d", CARD: "#2c3142",
    ACCENT: "#7986cb", ACCENT_DARKER: "#3f51b5", RED: "#f44336", BLUE: "#90caf9",
    YELLOW: "#ffeb3b", DISABLED: "#607d8b", ACC_ICON_BG: "#373e52", DLG_BG: "#2d3347",
    DLG_FG: "#c5cae9", DLG_CARD: "#353c50", TV_HEAD_BG: "#242938", TV_ODD: "#2c3142",
    TV_EVEN: "#30384d", SB_TROUGH: "#2c3142", SB_BG: "#242938", BTN_FG: "#ffffff",
    CB_LIST_BG: "#353c50", CB_LIST_FG: "#c5cae9", CB_LIST_SEL_BG: "#7986cb",
    CB_LIST_SEL_FG: "#ffffff", LBL_DESC_FG: "#9e9e9e",
}

# IVO+GPT (IVO-HEXCODE SELECTION, GPT-HEXCODE REPETITIONS)
THEME_ECO_LIGHT = {
    BG: "#e8eaf6", FG: "#303f9f", SIDEBAR: "#c5cae9", CARD: "#ffffff",
    ACCENT: "#5c6bc0", ACCENT_DARKER: "#3949ab", RED: "#e53935", BLUE: "#90caf9",
    YELLOW: "#fdd835", DISABLED: "#90a4ae", ACC_ICON_BG: "#cfd8dc", DLG_BG: "#f5f5ff",
    DLG_FG: "#303f9f", DLG_CARD: "#ffffff", TV_HEAD_BG: "#e8eaf6", TV_ODD: "#ffffff",
    TV_EVEN: "#f5f5f5", SB_TROUGH: "#ffffff", SB_BG: "#e8eaf6", BTN_FG: "#ffffff",
    CB_LIST_BG: "#ffffff", CB_LIST_FG: "#303f9f", CB_LIST_SEL_BG: "#5c6bc0",
    CB_LIST_SEL_FG: "#ffffff", LBL_DESC_FG: "#757575",
}
# Global variable holding the currently active theme colors
theme_colors = THEME_ECO_DARK.copy()

# Fonts (Defined as constants)
FONT_FAMILY = "Segoe UI"
FONT_NORMAL = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_LARGE = (FONT_FAMILY, 14, "bold")
FONT_XLARGE = (FONT_FAMILY, 18, "bold")
FONT_XXLARGE = (FONT_FAMILY, 24, "bold")
FONT_ACCOUNT_NAME = (FONT_FAMILY, 12)
FONT_TITLE = (FONT_FAMILY, 12, "bold")
FONT_SMALL = (FONT_FAMILY, 8)
FONT_DESC = (FONT_FAMILY, 9)

# --- Data Store ---
# Central dictionary holding application state and user data
# Using a more descriptive name for the main data dictionary
app_state = {
    "user_profiles": {},        # Dict: {user_id: {"name": "...", "icon_color": "#..."}}
    "current_user_id": None,    # ID of the currently logged-in user
    "emission_factors": {},     # Dict: {factor_id: float_value} loaded from CSV/defaults
    "activities": [],           # List of dicts: [{"timestamp": ..., "category": ..., "details": {...}, "carbon_footprint": ...}]
    "activity_log": [],         # List of dicts: [{"timestamp": ..., "action": "..."}] for user actions
    "settings": {               # User-specific settings
        "theme": "eco_dark",    # Default theme
        "conversion": "CO2e"    # Default display unit
    },
    "categories": {},           # Loaded categories (from BASE_CATEGORIES)
}

# --- Core Categories --
CategoryType = Dict[str, Any]  # More general: inner dicts can contain different types
# Or, if the inner dictionaries REALLY only contain strings, use this:
# CategoryType = Dict[str, Dict[str, str]]
BASE_CATEGORIES: Dict[str, CategoryType] = {
    "residential": {"name": "Residential", "icon": "🏠", "type": "category"},
    "travel": {"name": "Travel", "icon": "🚗", "type": "category"},
    "food": {"name": "Food", "icon": "🍎", "type": "category"},
    "shopping": {"name": "Goods & Waste", "icon": "🛍", "type": "category"},
    "services": {"name": "Services", "icon": "💼", "type": "category"},
    "digital": {"name": "Digital", "icon": "💻", "type": "category"},
}
# --- File Paths & Constants ---
DATA_DIR = "EcoHubData" # Folder to store all user data and configurations
USER_PROFILES_CSV = os.path.join(DATA_DIR, "user_profiles.csv")
EMISSION_FACTORS_CSV = os.path.join(DATA_DIR, "emission_factors.csv")
# Colors for user profile icons on the selection screen
ACCOUNT_ICON_COLORS = ["#8BC34A", "#4CAF50", "#66BB6A", "#9CCC65", "#AED581", "#C5E1A5", "#DCEDC8", "#E8F5E9"]
MAX_ACTIVITY_LOG_SIZE = 150 # Maximum user actions in history
# PHP/USD Conversion
PHP_TO_USD_RATE = 57

# --- Default Emission Factors ---
# Stored as a constant dictionary
DEFAULT_EMISSION_FACTORS = {
    # --- Residential ---
    "res_elec_usage_ph_nat_avg_kwh": 0.6032, "res_heat_nat_gas_therm": 5.3,
    "res_heat_heating_oil_gallon": 1.01, "res_heat_propane_gallon": 0.863,
    "res_heat_wood_hardwood_cord": 2208.93, "res_heat_wood_softwood_cord": 2292.60,
    "res_water_elec_kwh": 0.6032, "res_water_gas_therm": 5.3, "res_water_solar_thermal_kwh": 0.041,
    "res_renew_solar_panels_kwh": -0.041, "res_renew_wind_turbines_kwh": -0.011,
    # --- Transportation ---
    "trans_pv_gasoline_km": 0.195, "trans_pv_diesel_km": 0.224, "trans_pv_electric_km": 0.104,
    "trans_pub_bus_pkm": 0.093, "trans_pub_train_pkm": 0.058, "trans_pub_subway_pkm": 0.031,
    "trans_pub_jeepney_pkm": 0.114, "trans_pub_motorcycle_pkm": 0.103,
    "trans_air_short_pkm": 0.251, "trans_air_medium_pkm": 0.209, "trans_air_long_pkm": 0.195,
    "trans_air_cabin_economy": 1.0, "trans_air_cabin_business": 3.0, "trans_air_cabin_first": 9.0,
    # --- Food Consumption ---
    "food_prod_beef_kg_kg": 60.0, "food_prod_lamb_kg_kg": 24.0, "food_prod_pork_kg_kg": 7.0,
    "food_prod_poultry_kg_kg": 6.0, "food_prod_seafood_kg_kg": 5.0, "food_prod_dairy_kg_kg": 3.0,
    "food_prod_eggs_kg_kg": 4.5, "food_prod_vegetables_kg_kg": 0.88, "food_prod_fruits_kg_kg": 0.46,
    "food_prod_grains_kg_kg": 1.4, "food_prod_legumes_kg_kg": 0.9,
    "food_miles_avg_km": 2400, "food_miles_truck_tkm": 0.07, "food_miles_train_tkm": 0.03, "food_miles_air_tkm": 1.5,
    "food_proc_meat_tonne": 0.3, "food_proc_bread_tonne": 8.0, "food_proc_sugar_tonne": 10.0, "food_proc_fats_tonne": 10.0,
    "food_proc_cereals_tonne": 1.0, "food_proc_animalfeed_tonne": 1.0, "food_proc_coffee_tonne": 0.55,
    "food_pack_paper_kg_kg": 1.25, "food_pack_plastic_kg_kg": 3.25, "food_pack_glass_kg_kg": 1.5,
    "food_pack_recycled_kg_kg": 3.0, "food_pack_steel_kg_kg": 3.0, "food_pack_biopolymer_kg_kg": 3.0,
    "food_farm_fertilizer_conventional_kgN": 2.98, "food_farm_fertilizer_organic_kgN": 2.4, "food_farm_avg_N_per_kg_crop": 0.02,
    "food_region_luzon_kg_crop": 0.5, "food_region_visayas_kg_crop": 0.55, "food_region_mindanao_kg_crop": 0.5,
    # --- Goods Consumption and Waste ---
    "spending_clothing_usd": 0.3, "spending_electronics_usd": 0.3, "spending_appliances_usd": 0.3,
    "spending_furniture_usd": 0.3, "spending_other_goods_usd": 0.3,
    "goods_region_urban_retail_mult": 1.6666667, "goods_region_rural_retail_mult": 1.3333333, # Precalculated
    "waste_landfill_low_ch4_kg_kg": 0.42, "waste_landfill_med_ch4_kg_kg": 1.90, "waste_landfill_high_ch4_kg_kg": 2.52,
    "waste_incineration_kg_kg": 1.20, "waste_region_urban_landfill_kg_kg": 1.90, "waste_region_rural_landfill_kg_kg": 1.40,
    "waste_recycle_paper_kg_kg": -0.46, "waste_recycle_plastic_kg_kg": -1.08, "waste_recycle_glass_kg_kg": -0.31,
    "waste_recycle_aluminum_kg_kg": -8.14, "waste_recycle_steel_kg_kg": -0.86, "waste_recycle_copper_kg_kg": -3.0,
    "waste_recycle_avg_mix_kg_kg": -0.80,
    # --- Service Industries ---
    "serv_drycleaning_base_kg_garment": 2.0, "serv_drycleaning_region_urban_kg_garment": 7.0, "serv_drycleaning_region_rural_kg_garment": 7.0,
    "serv_landscaping_base_m2": 0.2, "serv_landscaping_region_urban_m2": 0.2, "serv_landscaping_region_rural_m2": 0.15,
    # --- Digital Footprint ---
    "digital_grid_luzon_kwh": 0.85, "digital_grid_visayas_kwh": 0.85, "digital_grid_mindanao_kwh": 0.82,
    "digital_grid_default_kwh": 0.6032,
    "digital_datacenter_kwh_gb": 5.12, "digital_network_kwh_gb": 0.06,
    "digital_laptop_kwh_hour": 0.055 / 24.0, "digital_mobile_kwh_hour": 7.3 / (365.0 * 24.0), "digital_tablet_kwh_hour": 0.06 / 24.0,
    "digital_stream_low_kwh_hour": 0.1, "digital_stream_medium_kwh_hour": 0.2, "digital_stream_high_kwh_hour": 0.7,
    "digital_game_low_kwh_hour": 0.14, "digital_game_high_kwh_hour": 0.35,
}

# --- Utility Functions ---

# EXPENSEWISE
def create_stylish_button(parent, text, command, style="TButton", **kwargs):
    """Creates a ttk.Button with specified style."""
    return ttk.Button(parent, text=text, command=command, style=style, **kwargs)

# EXPENSEWISE
def create_card_frame(parent):
    """Creates a standard styled frame for holding card content."""
    # Use the theme_colors dictionary directly
    return tk.Frame(parent, bg=theme_colors[CARD], relief=tk.FLAT, bd=0)

# GUTIERREZ+KATSUYA
def format_carbon_emission(amount_kg_co2e, conversion_unit="CO2e"):
    """Formats a carbon footprint value (in kg CO2e) into the desired display unit."""
    if amount_kg_co2e is None: return "N/A"
    try:
        # Ensure amount is float for calculations
        amount_kg_co2e = float(amount_kg_co2e)
        # Conversion factors (refine as needed)
        trees_absorption_kg_per_year = 21.7 # Avg kg CO2 per urban tree per year
        car_emissions_kg_per_year = 4600   # Avg US passenger vehicle kg CO2e per year (EPA)

        if conversion_unit == "Trees (Absorbed CO2 per Year)":
            if abs(amount_kg_co2e) < 0.001: # Handle near-zero values
                converted_value = 0.0
            elif trees_absorption_kg_per_year == 0: # Avoid division by zero
                logging.warning("Tree absorption factor is zero.")
                return "N/A (Config Error)"
            else:
                 converted_value = amount_kg_co2e / trees_absorption_kg_per_year
            unit_label = "Trees/yr"
            # Adjust precision based on magnitude
            precision = 2 if abs(converted_value) >= 0.1 else 3
            return f"{converted_value:,.{precision}f} {unit_label}"

        elif conversion_unit == "Cars (Emitted CO2 per Year)":
            if abs(amount_kg_co2e) < 0.001:
                converted_value = 0.0
            elif car_emissions_kg_per_year == 0:
                logging.warning("Car emissions factor is zero.")
                return "N/A (Config Error)"
            else:
                converted_value = amount_kg_co2e / car_emissions_kg_per_year
            unit_label = "Cars/yr"
            return f"{converted_value:,.4f} {unit_label}" # Always show reasonable precision

        else: # Default to CO2e
            return f"{amount_kg_co2e:,.2f} kg CO₂e"

    except (ValueError, TypeError) as e:
        logging.warning(f"Invalid value for carbon formatting: {amount_kg_co2e} ({type(amount_kg_co2e)}), Error: {e}")
        return "Invalid"

# EXPENSEWISE
def log_activity(action):
    """Logs a user action to the in-memory activity log."""
    user_id = app_state.get("current_user_id")
    if not user_id: return # Don't log if no user context
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {"timestamp": timestamp, "action": action}

    # Ensure activity_log is a list and append
    if not isinstance(app_state.get("activity_log"), list):
        app_state["activity_log"] = []
    app_state["activity_log"].append(log_entry)

    # Trim the log if it exceeds the maximum size
    if len(app_state["activity_log"]) > MAX_ACTIVITY_LOG_SIZE:
        # Use slicing for potentially faster removal from the front
        app_state["activity_log"] = app_state["activity_log"][-MAX_ACTIVITY_LOG_SIZE:]
        # Alternative: app_state["activity_log"].pop(0) is also fine for moderate sizes

# EXPENSEWISE
def get_unique_id(prefix):
    """Generates a reasonably unique ID string."""
    timestamp = int(datetime.datetime.now().timestamp() * 1000) # Milliseconds for finer granularity
    random_part = random.randint(10000, 99999)
    return f"{prefix}_{timestamp}_{random_part}"

# IVO-ONLY
def ensure_data_dir():
    """Creates the data directory if it doesn't exist."""
    if not os.path.exists(DATA_DIR):
        try:
            os.makedirs(DATA_DIR)
            logging.info(f"Created data directory: {DATA_DIR}")
            # If directory was just created, create default factors file
            if not os.path.exists(EMISSION_FACTORS_CSV):
                create_default_emission_factors_csv()
        except OSError as e:
            logging.critical(f"Could not create data directory '{DATA_DIR}': {e}")
            messagebox.showerror("Directory Error", f"Could not create data directory '{DATA_DIR}':\n{e}\nApplication cannot continue.")
            exit(1) # Critical error

# IVO-ONLY
def create_default_emission_factors_csv():
    """Creates the emission_factors.csv file with defaults if it doesn't exist."""
    if os.path.exists(EMISSION_FACTORS_CSV):
         logging.debug(f"Emission factors file already exists: {EMISSION_FACTORS_CSV}")
         return
    try:
        with open(EMISSION_FACTORS_CSV, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['factor_id', 'value', 'unit', 'category', 'description', 'source_notes']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            logging.info(f"Writing {len(DEFAULT_EMISSION_FACTORS)} default factors to {EMISSION_FACTORS_CSV}")

            for factor_id, value in DEFAULT_EMISSION_FACTORS.items():
                # Determine Unit, Category, Description from factor_id
                unit = "kg CO2e/unit" # Default
                category = "Unknown"
                description = factor_id.replace("_", " ").title() # Basic description
                source_notes = ""

                # Simplified category determination
                cat_prefix = factor_id.split('_')[0]
                category_map = {
                    "res": "Residential", "trans": "Transportation", "food": "Food",
                    "goods": "Goods", "waste": "Waste", "serv": "Services",
                    "spending": "Goods", "digital": "Digital"
                }
                category = category_map.get(cat_prefix, "Unknown")

                # Determine Unit Suffix
                id_lower = factor_id.lower()
                if "_kwh_hour" in id_lower: unit = "kWh/hour"
                elif "_kwh_gb" in id_lower: unit = "kWh/GB"
                elif "kwh" in id_lower: unit = "kg CO2e/kWh"
                elif "_therm" in id_lower: unit = "kg CO2e/therm"
                elif "_gallon" in id_lower: unit = "kg CO2e/gallon"
                elif "_cord" in id_lower: unit = "kg CO2e/cord"
                elif "_pkm" in id_lower: unit = "kg CO2e/pkm"
                elif "_tkm" in id_lower: unit = "kg CO2e/tonne-km"
                elif "_km" in id_lower: unit = "kg CO2e/km"
                elif "_kg_kg" in id_lower: unit = "kg CO2e/kg"
                elif "_tonne" in id_lower: unit = "kg CO2e/tonne"
                elif "_m2" in id_lower: unit = "kg CO2e/m²"
                elif "_usd" in id_lower or factor_id.startswith("spending_"): unit = "kg CO2e/PHP (Approx)"; description = description.replace("Usd", "per PHP")
                elif "_gb" in id_lower: unit = "kg CO2e/GB"
                elif "_kgN" in id_lower: unit = "kg CO2e/kg Nitrogen"
                elif "_kg_crop" in id_lower: unit = "kg CO2e/kg Crop"
                elif "_mult" in id_lower: unit = "Multiplier"
                elif "cabin_" in id_lower: unit = "Multiplier"
                elif "_hour" in id_lower: unit = "kWh/hour" # Catch device hours without kWh prefix

                # Specific Overrides
                if factor_id == "food_miles_avg_km": unit = "km"; description = "Average Food Travel Distance (Info)"
                if factor_id.startswith("res_renew_") or factor_id.startswith("waste_recycle_"): description += " (Saving)"
                if "nat_avg" in factor_id: description += " (Nat Avg)"
                if "region_" in factor_id: description += " (Region Adj)"
                if factor_id == "spending_other_goods_usd": description = "Spending Other Goods per PHP (Approx)"

                # Clean up description prefix
                desc_cleaned = description.replace(category + " ", "") if category != "Unknown" else description

                writer.writerow({
                    'factor_id': factor_id, 'value': value, 'unit': unit,
                    'category': category, 'description': desc_cleaned.strip(),
                    'source_notes': source_notes
                })
        logging.info(f"Successfully created default emission factors file: {EMISSION_FACTORS_CSV}")
    except IOError as e:
        logging.error(f"Could not write default emission factors file '{EMISSION_FACTORS_CSV}': {e}")
        messagebox.showerror("File Error", f"Could not write default emission factors:\n{e}", parent=None) # No parent context here
    except Exception as e:
         logging.exception(f"Unexpected error creating default factors file: {e}")
         messagebox.showerror("File Error", f"Could not write default emission factors:\n{e}", parent=None)

# --- Data Loading/Saving ---

# EXPENSEWISE
# User Profiles (CSV)
def load_user_profiles_from_csv():
    """Loads user profile data from CSV, ensuring robustness."""
    ensure_data_dir()
    profiles = {}
    created_demo = False
    required_fields = ['user_id', 'name', 'icon_color']

    if not os.path.exists(USER_PROFILES_CSV):
        logging.warning(f"'{USER_PROFILES_CSV}' not found. Creating demo profile.")
    else:
        try:
            with open(USER_PROFILES_CSV, mode='r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                # Strict check for required fields
                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_fields):
                    raise ValueError(f"CSV header missing required columns: {required_fields}")

                for row_num, row in enumerate(reader, 1):
                    try:
                        user_id = row.get('user_id', '').strip()
                        name = row.get('name', '').strip()
                        icon_color = row.get('icon_color', '').strip()

                        if not user_id or not name:
                            logging.warning(f"Skipping invalid row {row_num} in profiles CSV (missing ID or Name): {row}")
                            continue

                        # Validate or assign color
                        if not icon_color or not (icon_color.startswith('#') and len(icon_color) == 7):
                            logging.warning(f"Invalid/missing icon_color '{icon_color}' for user {name} (row {row_num}), assigning random.")
                            icon_color = random.choice(ACCOUNT_ICON_COLORS)

                        if user_id in profiles:
                             logging.warning(f"Duplicate user ID '{user_id}' found in profiles CSV (row {row_num}). Overwriting.")
                        profiles[user_id] = {"name": name, "icon_color": icon_color}

                    except Exception as row_e:
                        logging.error(f"Error processing profile row {row_num} ({row}): {row_e}. Skipping.")
                        continue # Skip problematic rows

        except (ValueError, csv.Error, IOError) as e:
            logging.exception(f"Failed to load user profiles from '{USER_PROFILES_CSV}': {e}. Attempting demo profile.")
            profiles.clear() # Clear potentially corrupt data before creating demo
        except Exception as e:
             logging.exception(f"Unexpected error loading user profiles: {e}")
             profiles.clear()

    # If no valid profiles loaded after trying the file, create a demo user
    if not profiles:
        logging.warning("No valid user profiles loaded. Creating demo user.")
        demo_id = get_unique_id("user_demo")
        profiles[demo_id] = {"name": "Eco User", "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
        created_demo = True

    app_state["user_profiles"] = profiles
    if created_demo:
        save_user_profiles_to_csv() # Save the newly created demo profile

# EXPENSEWISE
def save_user_profiles_to_csv():
    """Saves the current user profiles from app_state to the CSV file.

    Returns:
        bool: True if save was successful, False otherwise.
    """
    try:
        ensure_data_dir() # Ensure directory exists first
    except SystemExit: # Catch exit if ensure_data_dir failed critically
        logging.critical("Cannot save profiles, data directory error.")
        return False

    profiles_to_save = app_state.get("user_profiles")
    required_fields = ['user_id', 'name', 'icon_color']

    if not profiles_to_save or not isinstance(profiles_to_save, dict):
        logging.warning("No valid user profiles found in app_state to save. Writing header only.")
        try:
            with open(USER_PROFILES_CSV, mode='w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=required_fields)
                writer.writeheader()
            logging.info(f"Created empty user profiles file with header: '{USER_PROFILES_CSV}'.")
            return True # Writing header successfully is considered success
        except IOError as e:
            logging.error(f"Could not write header to empty '{USER_PROFILES_CSV}': {e}")
            return False # Indicate failure

    # Proceed with saving actual profiles
    try:
        with open(USER_PROFILES_CSV, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=required_fields, extrasaction='ignore')
            writer.writeheader()
            for user_id, details in profiles_to_save.items():
                if not isinstance(details, dict):
                    logging.warning(f"Skipping saving invalid profile data for ID {user_id}: {details}")
                    continue
                row_data = {
                    'user_id': user_id,
                    'name': details.get('name', f'Unnamed_{user_id}'),
                    'icon_color': details.get('icon_color', random.choice(ACCOUNT_ICON_COLORS))
                }
                writer.writerow(row_data)
        logging.info(f"User profiles saved successfully to '{USER_PROFILES_CSV}'.")
        return True # Explicitly return True on success
    except IOError as e:
        logging.error(f"IOError saving user profiles to '{USER_PROFILES_CSV}': {e}")
        return False # Indicate failure
    except Exception as e:
        logging.exception(f"Unexpected error saving user profiles to '{USER_PROFILES_CSV}'")
        return False # Indicate failure

# IVO-ONLY
# User Data (JSON Helpers)
def get_user_data_file_path(user_id, data_type):
    """Constructs the file path for a user's data file (JSON)."""
    base_filename = f"{user_id}_{data_type}.json"
    return os.path.join(DATA_DIR, base_filename)

# IVO-ONLY
def _load_json_data(file_path, default_value_factory=dict):
    """Loads data from JSON, returning a default value if file missing/invalid."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            logging.debug(f"Loaded JSON: {file_path}")
            return data
    except FileNotFoundError:
        logging.info(f"JSON file not found: {file_path}. Using default.")
        return default_value_factory() # Call factory to get new default instance
    except json.JSONDecodeError as e:
        logging.error(f"Error decoding JSON {file_path}: {e}. Using default.")
        return default_value_factory()
    except Exception as e:
        logging.exception(f"Unexpected error loading JSON {file_path}: {e}")
        return default_value_factory()

# IVO-ONLY
def _save_json_data(file_path, data):
    """Saves data to a JSON file."""
    try:
        # Ensure the directory exists right before saving
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        logging.debug(f"Saved JSON: {file_path}")
        return True
    except (IOError, TypeError) as e: # Catch specific expected errors
        logging.error(f"Error saving JSON to {file_path}: {e}")
        return False
    except Exception as e:
        logging.exception(f"Unexpected error saving JSON to {file_path}: {e}")
        return False

# IVO-ONLY
# Emission Factors (CSV)
def load_emission_factors():
    """Loads factors from CSV, merging with defaults."""
    ensure_data_dir()
    factors = {}
    required_fields = ['factor_id', 'value']

    try:
        # Attempt to load from CSV if it exists
        if os.path.exists(EMISSION_FACTORS_CSV):
            with open(EMISSION_FACTORS_CSV, mode='r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_fields):
                     logging.warning(f"{EMISSION_FACTORS_CSV} header missing required columns. Merging with defaults.")
                     # Continue to load what we can, defaults will fill gaps
                else:
                    for row_num, row in enumerate(reader, 1):
                        try:
                            factor_id = row.get('factor_id', '').strip()
                            value_str = row.get('value', '').strip()
                            if not factor_id or value_str is None: continue # Skip empty IDs or None values

                            try:
                                factors[factor_id] = float(value_str)
                            except (ValueError, TypeError):
                                logging.warning(f"Invalid value '{value_str}' for factor '{factor_id}' in CSV row {row_num}. Skipping.")
                        except Exception as row_e:
                            logging.error(f"Error processing factor row {row_num}: {row_e}. Skipping.")
            logging.info(f"Loaded {len(factors)} factors from {EMISSION_FACTORS_CSV}.")
        else:
             raise FileNotFoundError

    except FileNotFoundError:
        logging.warning(f"Emission factors file '{EMISSION_FACTORS_CSV}' not found. Creating and using defaults.")
        create_default_emission_factors_csv()
        factors = DEFAULT_EMISSION_FACTORS.copy()
    except (csv.Error, IOError, ValueError) as e:
        logging.warning(f"Failed to load/process factors from '{EMISSION_FACTORS_CSV}': {e}. Merging with defaults.")
    except Exception as e:
        logging.exception(f"Unexpected error loading factors: {e}. Merging with defaults.")
        # Continue, defaults will merge below

    # Merge loaded factors with defaults (defaults overwrite if missing/invalid in CSV)
    # Start with a copy of defaults, then update with valid loaded factors
    final_factors = DEFAULT_EMISSION_FACTORS.copy()
    final_factors.update(factors)

    # Final check if factors dictionary is somehow empty
    if not final_factors:
         logging.critical("Emission factors are empty after all loading attempts. Using fallback defaults.")
         final_factors = DEFAULT_EMISSION_FACTORS.copy()

    app_state["emission_factors"] = final_factors
    logging.info(f"Final emission factor count: {len(app_state['emission_factors'])}")

# IVO-ONLY
# Combined User Data Loading/Saving
def load_user_data(user_id):
    """Loads all data (settings, activities, logs, factors) for a user."""
    logging.info(f"Loading data for user: {user_id}")
    app_state["current_user_id"] = user_id

    # 1. Load/Ensure Emission Factors (shared, load once per session effectively)
    if not app_state.get("emission_factors"): # Only load if not already loaded
        load_emission_factors()

    # 2. Load User-Specific Data (Settings, Activities, Log)
    user_data_config = {
        "settings": {"default_factory": lambda: {"theme": "eco_dark", "conversion": "CO2e"}},
        "activities": {"default_factory": list},
        "activity_log": {"default_factory": list},
    }

    for key, config in user_data_config.items():
        file_path = get_user_data_file_path(user_id, key)
        loaded_data = _load_json_data(file_path, default_value_factory=config["default_factory"])

        # --- Post-load validation and cleanup ---
        default_instance = config["default_factory"]() # Get a default instance for comparison/fallback
        if key == "settings":
            if not isinstance(loaded_data, dict):
                logging.warning(f"Settings data for {user_id} invalid, using defaults.")
                loaded_data = default_instance
            # Ensure required keys exist using defaults
            for k, v in default_instance.items():
                loaded_data.setdefault(k, v)

        elif key == "activities":
            if not isinstance(loaded_data, list):
                logging.warning(f"Activities data for {user_id} invalid, using empty list.")
                loaded_data = default_instance
            # Validate each activity structure
            valid_activities = []
            for i, activity in enumerate(loaded_data):
                 if isinstance(activity, dict) and \
                    activity.get("timestamp") and \
                    activity.get("category") and \
                    isinstance(activity.get("activity_details"), dict): # Ensure details is dict
                    activity.setdefault("carbon_footprint", None) # Ensure key exists
                    valid_activities.append(activity)
                 else:
                     logging.warning(f"Skipping invalid activity record #{i+1} for user {user_id}: {activity}")
            loaded_data = valid_activities

        elif key == "activity_log":
             if not isinstance(loaded_data, list):
                 logging.warning(f"Activity log data for {user_id} invalid, using empty list.")
                 loaded_data = default_instance
             # Optional: Validate log entry format here if needed

        app_state[key] = loaded_data # Store validated data

    # 3. Set base categories (always static)
    app_state["categories"] = BASE_CATEGORIES

    logging.info(f"Data loading finished for {user_id}. Theme: {app_state['settings']['theme']}, Activities: {len(app_state['activities'])}")

# IVO-ONLY
def save_user_data(user_id):
    """Saves user-specific data (settings, activities, logs) to JSON files."""
    if not user_id:
        logging.error("Attempted to save data without a valid user ID.")
        return False

    logging.info(f"Saving data for user: {user_id}")
    ensure_data_dir()

    user_data_keys = ["settings", "activities", "activity_log"]
    save_success_overall = True

    for key in user_data_keys:
        file_path = get_user_data_file_path(user_id, key)
        data_to_save = app_state.get(key)

        if data_to_save is None: # Should not happen if load_user_data ran correctly
            logging.warning(f"No data for '{key}' found for user {user_id}. Skipping save.")
            continue

        if not _save_json_data(file_path, data_to_save):
            save_success_overall = False
            logging.error(f"FAILED to save '{key}' to '{file_path}'.")
            # Show error message ONLY if overall save fails later

    if not save_success_overall:
        logging.error(f"One or more data files failed to save for user: {user_id}.")
        messagebox.showerror("Save Error", f"Failed to save some user data for {user_id}. Please check logs.")

    return save_success_overall

# --- Accounts Page Class (Profile Selection) ---
class AccountsPage(tk.Tk):
    """Initial window for selecting or creating a user profile."""

    # EXPENSEWISE
    def __init__(self):
        super().__init__()
        self.selected_user_id = None # Store the ID of the selected user
        # Apply theme directly (Accounts page always uses dark theme)
        self.configure(bg=THEME_ECO_DARK[BG])
        self.title("ECOHUB - Select Profile")
        self.geometry("800x450") # Slightly taller for padding

        # Ensure data dir exists and load profiles
        ensure_data_dir()
        load_user_profiles_from_csv()

        # --- Styling ---
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        # Removed Account.TButton styles as profile selection uses Frame+Label now
        self.style.configure('AddAccount.TButton', background=THEME_ECO_DARK[BG], foreground=THEME_ECO_DARK[DISABLED], font=FONT_XXLARGE, borderwidth=1, relief=tk.SOLID, bordercolor=THEME_ECO_DARK[DISABLED])
        self.style.map('AddAccount.TButton', foreground=[('active', THEME_ECO_DARK[FG])], bordercolor=[('active', THEME_ECO_DARK[FG])])
        self.style.configure('Exit.TButton', font=FONT_NORMAL, foreground=THEME_ECO_DARK[FG], background="#555555", borderwidth=1, relief=tk.SOLID)
        self.style.map('Exit.TButton', background=[('active', THEME_ECO_DARK[RED])], foreground=[('active', THEME_ECO_DARK[BTN_FG])])

        # --- Layout ---
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Title
        title_label = tk.Label(self, text="Who's Tracking?", font=FONT_XLARGE, bg=THEME_ECO_DARK[BG], fg=THEME_ECO_DARK[FG])
        title_label.grid(row=0, column=0, pady=(40, 20)) # Adjusted padding

        # Accounts Frame (Scrollable potentially, but usually few profiles)
        self.accounts_frame = tk.Frame(self, bg=THEME_ECO_DARK[BG])
        self.accounts_frame.grid(row=1, column=0, sticky="n", pady=10) # Reduced bottom padding
        self.accounts_frame.grid_columnconfigure(0, weight=1) # Center the inner frame

        self.display_user_profiles()

        # Exit Button
        exit_button = ttk.Button(self, text="Exit Application", command=self.exit_app, style="Exit.TButton", width=15)
        exit_button.place(relx=0.98, rely=0.98, anchor='se', x=-15, y=-15) # Adjusted position

        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.exit_app) # Handle window close button

    # EXPENSEWISE
    def center_window(self):
        """Centers the window on the screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x_pos = (self.winfo_screenwidth() // 2) - (width // 2)
        y_pos = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x_pos}+{y_pos}')
        self.attributes('-alpha', 1.0) # Ensure fully opaque

    # EXPENSEWISE
    def display_user_profiles(self):
        """Creates and displays the profile icons and add button."""
        # Clear previous widgets
        for widget in self.accounts_frame.winfo_children():
            widget.destroy()

        # Use an inner frame to easily center the grid of profiles
        inner_frame = tk.Frame(self.accounts_frame, bg=THEME_ECO_DARK[BG])
        inner_frame.pack(pady=10) # Pack centers it horizontally due to accounts_frame config

        max_cols = 4 # Max profiles per row
        profiles = app_state.get("user_profiles", {})
        # Sort profiles alphabetically by name for consistent order
        sorted_profiles = sorted(profiles.items(), key=lambda item: item[1].get('name', '').lower())

        col_count = 0
        row_count = 0
        for user_id, details in sorted_profiles:
            profile_col = col_count % max_cols
            profile_row = row_count + (col_count // max_cols)

            container = tk.Frame(inner_frame, bg=THEME_ECO_DARK[BG])
            container.grid(row=profile_row, column=profile_col, padx=25, pady=15, sticky="n") # Increased padding

            name = details.get('name', 'Unknown')
            initial = name[0].upper() if name else "?"
            icon_color = details.get('icon_color', random.choice(ACCOUNT_ICON_COLORS))

            # Clickable Frame for Icon
            icon_frame = tk.Frame(container, bg=icon_color, width=100, height=100, cursor="hand2")
            icon_frame.pack(pady=(0, 8)) # Space between icon and name
            icon_frame.pack_propagate(False) # Prevent label from shrinking frame

            initial_label = tk.Label(icon_frame, text=initial, font=FONT_XXLARGE, bg=icon_color, fg=THEME_ECO_DARK[BTN_FG])
            initial_label.place(relx=0.5, rely=0.5, anchor="center")

            name_label = tk.Label(container, text=name, font=FONT_ACCOUNT_NAME, bg=THEME_ECO_DARK[BG], fg=THEME_ECO_DARK[FG], wraplength=110) # Wrap long names
            name_label.pack()

            # Bind click event to the container, icon frame, and labels
            # Use lambda to capture the correct user_id for each profile
            for widget in [container, icon_frame, initial_label, name_label]:
                widget.bind("<Button-1>", lambda e, u_id=user_id: self.select_user(u_id))

            col_count += 1

        # Add "+" button for creating new profiles
        add_col = col_count % max_cols
        add_row = row_count + (col_count // max_cols)
        add_container = tk.Frame(inner_frame, bg=THEME_ECO_DARK[BG])
        add_container.grid(row=add_row, column=add_col, padx=25, pady=15, sticky="n")

        # Use a Frame with border for the add icon visual
        add_icon_frame = tk.Frame(add_container, bg=THEME_ECO_DARK[BG], width=100, height=100, cursor="hand2",
                                  relief=tk.SOLID, borderwidth=2,
                                  highlightbackground=THEME_ECO_DARK[DISABLED], # Use highlight for border
                                  highlightthickness=2, highlightcolor=THEME_ECO_DARK[DISABLED]) # Fallback color
        add_icon_frame.pack(pady=(0, 8))
        add_icon_frame.pack_propagate(False)

        add_label = tk.Label(add_icon_frame, text="+", font=FONT_XXLARGE, bg=THEME_ECO_DARK[BG], fg=THEME_ECO_DARK[DISABLED])
        add_label.place(relx=0.5, rely=0.5, anchor="center")
        add_label.bind("<Enter>", lambda e: add_label.config(fg=THEME_ECO_DARK[FG])) # Hover effect
        add_label.bind("<Leave>", lambda e: add_label.config(fg=THEME_ECO_DARK[DISABLED]))

        add_name_label = tk.Label(add_container, text="Add Profile", font=FONT_ACCOUNT_NAME, bg=THEME_ECO_DARK[BG], fg=THEME_ECO_DARK[DISABLED])
        add_name_label.pack()

        # Bind click event to add profile elements
        for widget in [add_container, add_icon_frame, add_label, add_name_label]:
            widget.bind("<Button-1>", self.add_user_profile_dialog)

    # EXPENSEWISE
    def add_user_profile_dialog(self, event=None):
        """Opens a dialog to get the new profile name."""
        # Configuration for the simple dialog
        fields = {"name": {"label": "Profile Name:", "required": True}}
        dialog = SimpleEntryDialog(self, "Create New Profile", fields)
        result = dialog.result # This blocks until the dialog is closed

        if result and isinstance(result, dict):
            name = result.get("name", "").strip()
            if not name:
                 messagebox.showerror("Invalid Input", "Profile name cannot be empty.", parent=self)
                 return

            # Check for duplicate names (case-insensitive)
            current_profiles = app_state.get("user_profiles", {})
            if any(prof.get('name', '').lower() == name.lower() for prof in current_profiles.values() if isinstance(prof, dict)):
                 messagebox.showerror("Duplicate Profile", f"A profile named '{name}' already exists.", parent=self)
                 return

            try:
                new_id = get_unique_id("user")
                new_profile = {"name": name, "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
                # Update in-memory state
                app_state["user_profiles"][new_id] = new_profile
                # Save to CSV
                if save_user_profiles_to_csv():
                    logging.info(f"Created and saved new profile: {name} ({new_id})")
                    self.display_user_profiles() # Refresh the display
                else:
                    # If save failed, revert the change in memory
                    app_state["user_profiles"].pop(new_id, None)
                    messagebox.showerror("Save Error", "Could not save the new profile. Please try again.", parent=self)

            except Exception as e:
                logging.exception("Error creating profile")
                messagebox.showerror("Error", f"Could not create profile.\n{e}", parent=self)

    # EXPENSEWISE
    def select_user(self, user_id):
        """Handles profile selection, sets the user ID, and closes the window."""
        profiles = app_state.get("user_profiles", {})
        if user_id in profiles:
            self.selected_user_id = user_id
            logging.info(f"Selected profile: {profiles[user_id].get('name', user_id)} ({user_id})")
            self.destroy() # Close the AccountsPage window
        else:
            logging.error(f"Attempted select non-existent user ID: {user_id}")
            messagebox.showerror("Error", "Selected profile not found. Reloading profiles.", parent=self)
            # Reload profiles from disk in case of inconsistency and refresh display
            load_user_profiles_from_csv()
            self.display_user_profiles()

    # EXPENSEWISE
    def exit_app(self):
        """Handles exit request from the Accounts Page."""
        logging.info("Exiting ECOHUB from Accounts Page.")
        self.selected_user_id = None # Ensure no user is selected if exiting
        self.destroy() # Close the window

# --- Main Application Class (ECOHUBApp) ---
class ECOHUBApp(tk.Tk):
    """The main application window after profile selection."""

    # IVO+GPT
    def __init__(self, user_id):
        super().__init__()
        self.current_user_id = user_id
        self._page_creation_lock = False # Prevent race conditions during page/theme switch
        self._full_exit_requested = False # Flag set by Settings->Exit Application

        # Load user data and apply initial theme
        load_user_data(self.current_user_id)
        self.current_theme = app_state.get("settings", {}).get("theme", "eco_dark")
        self._apply_theme_colors() # Sets the global theme_colors var

        self.title(f"ECOHUB - {app_state['user_profiles'].get(user_id, {}).get('name', 'Eco-User')}")
        self.geometry("1250x750") # Default size
        self.configure(bg=theme_colors[BG])

        # --- Styling ---
        self.style = ttk.Style(self)
        self._configure_styles() # Apply theme-dependent ttk styles

        # --- Layout ---
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1) # Main content area expands

        # Sidebar (Navigation)
        self.sidebar = Sidebar(self, self._show_page) # Pass reference to page switching method
        self.sidebar.grid(row=0, column=0, sticky="nsw")

        # Main Content Frame
        self.main_frame = tk.Frame(self, bg=theme_colors[BG])
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.current_page_frame = None # Holds the currently displayed page instance

        # Floating Action Button (FAB) for adding activities
        self.fab = create_stylish_button(self, "+", self.open_add_activity_dialog, style="FAB.TButton")
        self.fab.place(relx=0.98, rely=0.95, anchor='se') # Place in bottom-right

        # --- Initial State ---
        self._show_page("CarbonDashboardPage") # Show dashboard initially
        # Ensure the sidebar highlights the correct button
        if self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.highlight_button("CarbonDashboardPage")

        self.protocol("WM_DELETE_WINDOW", self.on_closing) # Handle window close button
        logging.info(f"ECOHUBApp initialized for user {user_id}.")

    # EXPENSEWISE
    def _apply_theme_colors(self):
        """Updates the global theme_colors dict based on current_theme."""
        global theme_colors
        if self.current_theme == "eco_light":
            theme_colors.update(THEME_ECO_LIGHT)
        else:
            theme_colors.update(THEME_ECO_DARK)

    # IVO+GPT
    def _configure_styles(self):
        """Configures all ttk styles based on the current theme."""
        self.style.theme_use('clam') # Base theme

        # General Widget Styling
        self.style.configure('.', background=theme_colors[BG], foreground=theme_colors[FG], font=FONT_NORMAL)
        self.style.configure('TFrame', background=theme_colors[BG])
        self.style.configure('TLabel', background=theme_colors[BG], foreground=theme_colors[FG], font=FONT_NORMAL)

        # Specific Label Styles
        self.style.configure('Sidebar.TLabel', background=theme_colors[SIDEBAR], foreground=theme_colors[FG])
        self.style.configure('Card.TLabel', background=theme_colors[CARD], foreground=theme_colors[FG])
        self.style.configure('Desc.TLabel', background=theme_colors[BG], foreground=theme_colors[LBL_DESC_FG], font=FONT_DESC)
        self.style.configure('CardDesc.TLabel', background=theme_colors[CARD], foreground=theme_colors[LBL_DESC_FG], font=FONT_DESC)
        self.style.configure('DialogDesc.TLabel', background=theme_colors[DLG_CARD], foreground=theme_colors[LBL_DESC_FG], font=FONT_DESC)
        self.style.configure('Title.TLabel', background=theme_colors[BG], foreground=theme_colors[FG], font=FONT_LARGE)
        self.style.configure('CardTitle.TLabel', background=theme_colors[CARD], foreground=theme_colors[FG], font=FONT_BOLD)
        self.style.configure('Accent.TLabel', background=theme_colors[CARD], foreground=theme_colors[ACCENT], font=FONT_BOLD)
        self.style.configure('Error.TLabel', background=theme_colors[BG], foreground=theme_colors[RED], font=FONT_BOLD)
        date_time_fg = theme_colors[DISABLED] if self.current_theme == 'eco_dark' else theme_colors[ACCENT_DARKER]
        self.style.configure('DateTime.TLabel', background=theme_colors[SIDEBAR], foreground=date_time_fg, font=FONT_SMALL)

        # Button Styling
        self.style.configure('TButton', background=theme_colors[ACCENT], foreground=theme_colors[BTN_FG], font=FONT_BOLD, padding=6, borderwidth=0, relief=tk.FLAT)
        self.style.map('TButton', background=[('active', theme_colors[ACCENT_DARKER])], foreground=[('active', theme_colors[BTN_FG])])
        self.style.configure('Sidebar.TButton', background=theme_colors[SIDEBAR], foreground=theme_colors[FG], font=FONT_BOLD, anchor='w', padding=(15, 8), borderwidth=0, relief=tk.FLAT)
        self.style.map('Sidebar.TButton', background=[('active', theme_colors[ACCENT]), ('selected', theme_colors[ACCENT])], foreground=[('active', theme_colors[BTN_FG]), ('selected', theme_colors[BTN_FG])])
        self.style.configure('FAB.TButton', background=theme_colors[ACCENT], foreground=theme_colors[BTN_FG], font=(FONT_FAMILY, 18, "bold"), padding=10, borderwidth=0, relief=tk.FLAT)
        self.style.map('FAB.TButton', background=[('active', theme_colors[ACCENT_DARKER])])

        # Treeview Styling
        self.style.configure("Treeview", background=theme_colors[CARD], foreground=theme_colors[FG], fieldbackground=theme_colors[CARD], rowheight=28, borderwidth=0, relief=tk.FLAT)
        self.style.configure("Treeview.Heading", background=theme_colors[TV_HEAD_BG], foreground=theme_colors[FG], font=FONT_BOLD, relief="flat", padding=(5, 5))
        self.style.map("Treeview.Heading", background=[('active', theme_colors[ACCENT])])
        self.style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

        # Progressbar Styling
        self.style.configure("TProgressbar", thickness=10, background=theme_colors[ACCENT], troughcolor=theme_colors[CARD])

        # Combobox Styling (Listbox part)
        # Use option_add for listbox customization (standard Tk practice)
        self.option_add('*TCombobox*Listbox*Background', theme_colors[CB_LIST_BG])
        self.option_add('*TCombobox*Listbox*Foreground', theme_colors[CB_LIST_FG])
        self.option_add('*TCombobox*Listbox*selectBackground', theme_colors[CB_LIST_SEL_BG])
        self.option_add('*TCombobox*Listbox*selectForeground', theme_colors[CB_LIST_SEL_FG])
        # Combobox widget itself
        self.style.configure('TCombobox', background=theme_colors[CARD], foreground=theme_colors[FG], fieldbackground=theme_colors[CARD], selectbackground=theme_colors[CARD], selectforeground=theme_colors[FG], arrowcolor=theme_colors[FG], borderwidth=1, padding=5, relief=tk.FLAT, bordercolor=theme_colors[DISABLED]) # Add subtle border
        self.style.map('TCombobox', fieldbackground=[('readonly', theme_colors[CARD])], bordercolor=[('focus', theme_colors[ACCENT])]) # Highlight border on focus

        # Entry Styling
        self.style.configure('TEntry', background=theme_colors[CARD], foreground=theme_colors[FG], fieldbackground=theme_colors[CARD], insertcolor=theme_colors[FG], borderwidth=1, padding=5, relief=tk.FLAT, bordercolor=theme_colors[DISABLED]) # Add subtle border
        self.style.map('TEntry', fieldbackground=[('focus', theme_colors[CARD])], bordercolor=[('focus', theme_colors[ACCENT])]) # Highlight border on focus

        # Notebook Styling
        self.style.configure('TNotebook', background=theme_colors[BG], borderwidth=0)
        self.style.configure('TNotebook.Tab', font=FONT_BOLD, padding=[10, 5], background=theme_colors[CARD], foreground=theme_colors[FG], borderwidth=0)
        self.style.map('TNotebook.Tab', background=[('selected', theme_colors[ACCENT])], foreground=[('selected', theme_colors[BTN_FG])])

        # Scrollbar Styling
        self.style.configure("Vertical.TScrollbar", background=theme_colors[SB_BG], troughcolor=theme_colors[SB_TROUGH], borderwidth=0, arrowcolor=theme_colors[FG], relief=tk.FLAT)
        self.style.map("Vertical.TScrollbar", background=[('active', theme_colors[ACCENT])])

        # Checkbutton/Radiobutton Styling
        # Base style matching background
        self.style.configure("TCheckbutton", background=theme_colors[BG], foreground=theme_colors[FG], font=FONT_NORMAL)
        self.style.map("TCheckbutton", indicatorcolor=[('selected', theme_colors[ACCENT])])
        self.style.configure("TRadiobutton", background=theme_colors[BG], foreground=theme_colors[FG], font=FONT_NORMAL)
        self.style.map("TRadiobutton", indicatorcolor=[('selected', theme_colors[ACCENT])])
        # Variations for placing on cards/dialog cards
        self.style.configure("Card.TRadiobutton", background=theme_colors[CARD], foreground=theme_colors[FG])
        self.style.map("Card.TRadiobutton", background=[('active', theme_colors[CARD])], indicatorcolor=[('selected', theme_colors[ACCENT])])
        self.style.configure("Dialog.TRadiobutton", background=theme_colors[DLG_CARD], foreground=theme_colors[DLG_FG])
        self.style.map("Dialog.TRadiobutton", background=[('active', theme_colors[DLG_CARD])], indicatorcolor=[('selected', theme_colors[ACCENT])])
        self.style.configure("Card.TCheckbutton", background=theme_colors[CARD], foreground=theme_colors[FG])
        self.style.map("Card.TCheckbutton", background=[('active', theme_colors[CARD])], indicatorcolor=[('selected', theme_colors[ACCENT])])
        self.style.configure("Dialog.TCheckbutton", background=theme_colors[DLG_CARD], foreground=theme_colors[DLG_FG])
        self.style.map("Dialog.TCheckbutton", background=[('active', theme_colors[DLG_CARD])], indicatorcolor=[('selected', theme_colors[ACCENT])])

        # Toolbutton (if used later)
        self.style.configure("Toolbutton", anchor="center", padding=5, font=FONT_NORMAL, background=theme_colors[CARD], foreground=theme_colors[FG], borderwidth=1, relief="raised")
        self.style.map("Toolbutton", relief=[('selected', 'sunken'), ('active', 'raised')], background=[('selected', theme_colors[ACCENT]), ('active', theme_colors[CARD])], foreground=[('selected', theme_colors[BTN_FG]), ('active', theme_colors[FG])])

        # --- Update non-ttk widget backgrounds ---
        self.configure(bg=theme_colors[BG])
        # Check existence before configuring existing widgets
        if hasattr(self, 'main_frame') and self.main_frame and self.main_frame.winfo_exists():
            self.main_frame.configure(bg=theme_colors[BG])
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.configure(bg=theme_colors[SIDEBAR])
            # Re-apply styles to sidebar widgets if needed (e.g., labels)
            self.sidebar.update_styles()
        if hasattr(self, 'fab') and self.fab and self.fab.winfo_exists():
            self.fab.configure(style="FAB.TButton") # Re-apply style

    # EXPENSEWISE
    def switch_theme(self, theme_name):
        """Switches the application theme and redraws UI elements."""
        if theme_name == self.current_theme or self._page_creation_lock:
            return # Avoid unnecessary switches or race conditions
        logging.info(f"Switching theme to: {theme_name}")
        self.current_theme = theme_name
        app_state["settings"]["theme"] = theme_name # Update setting in memory
        # Save the updated setting immediately
        # Do not proceed with UI changes if save fails
        if not save_user_data(self.current_user_id):
            # Error message shown by save_user_data
            # Revert theme choice in memory
            self.current_theme = "eco_light" if theme_name == "eco_dark" else "eco_dark"
            app_state["settings"]["theme"] = self.current_theme
            return

        self._apply_theme_colors() # Update global colors

        try:
            # Reconfigure all styles
            self._configure_styles()

            # Recreate Sidebar (simplest way to ensure all its styles update)
            current_page_name = "CarbonDashboardPage" # Default
            if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
                current_page_name = self.sidebar.get_current_page_name() or "CarbonDashboardPage"
                self.sidebar.stop_timer() # Essential before destroying
                self.sidebar.destroy()

            self.sidebar = Sidebar(self, self._show_page)
            self.sidebar.grid(row=0, column=0, sticky="nsw")

            # Recreate the current page with the new theme
            self._page_creation_lock = True # Prevent recursive calls
            try:
                self._show_page(current_page_name)
            finally:
                self._page_creation_lock = False # Release lock

            # Re-apply FAB style just in case
            if hasattr(self, 'fab') and self.fab and self.fab.winfo_exists():
                self.fab.configure(style="FAB.TButton")

            log_activity(f"Theme switched to {theme_name}")
        except Exception as e:
            logging.exception("Unexpected error during theme switch")
            messagebox.showerror("Theme Switch Error", f"Error switching theme:\n{e}", parent=self)
            # Attempt recovery by forcing dashboard display
            try:
                self._page_creation_lock = True
                self._show_page("CarbonDashboardPage")
            except Exception as recovery_e:
                 logging.error(f"Failed to recover after theme switch error: {recovery_e}")
            finally:
                self._page_creation_lock = False

    # IVO+GPT
    def _show_page(self, page_name):
        """Destroys the current page and displays the requested one."""
        if self._page_creation_lock: return # Prevent recursive calls during theme switch
        logging.info(f"Showing page: {page_name}")

        # Safely destroy the previous page frame
        if self.current_page_frame and self.current_page_frame.winfo_exists():
            try:
                # BasePage handles mousewheel unbinding in its destroy method
                self.current_page_frame.destroy()
            except Exception as e:
                logging.warning(f"Error destroying previous page frame '{type(self.current_page_frame).__name__}': {e}")
            finally:
                 self.current_page_frame = None

        # Map page name to class
        page_mapping = {
            "CarbonDashboardPage": CarbonDashboardPage,
            "ResidentialPage": ResidentialPage,
            "TravelPage": TravelPage,
            "FoodPage": FoodPage,
            "ShoppingPage": GoodsWastePage, # Map sidebar key to correct class
            "ServicesPage": ServicesPage,
            "DigitalPage": DigitalPage,
            "User History": UserHistoryPage,
            "Settings": SettingsPage
        }

        page_class = page_mapping.get(page_name)

        # Create and grid the new page
        if page_class:
            try:
                self.current_page_frame = page_class(self.main_frame, self) # Pass app instance
                self.current_page_frame.grid(row=0, column=0, sticky="nsew")
            except Exception as e:
                logging.exception(f"Error creating page '{page_name}'")
                messagebox.showerror("Page Load Error", f"Could not load page '{page_name}':\n{e}", parent=self)
                # Show a placeholder on error
                self.current_page_frame = PlaceholderPage(self.main_frame, f"Error loading {page_name}", self)
                self.current_page_frame.grid(row=0, column=0, sticky="nsew")
        else:
            logging.warning(f"Page class not found for '{page_name}'. Showing placeholder.")
            self.current_page_frame = PlaceholderPage(self.main_frame, page_name, self)
            self.current_page_frame.grid(row=0, column=0, sticky="nsew")

        # Highlight the corresponding button in the sidebar
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.highlight_button(page_name)
            self.sidebar.current_page_name = page_name # Keep track

    # IVO+GPT
    def open_add_activity_dialog(self):
        """Opens the modal dialog to add a new carbon activity."""
        # Dialog handles its own lifecycle (waits until closed)
        dialog = AddCarbonFootprintActivityDialog(self)
        # No return value needed here, dialog updates app_state directly

    # IVO+GPT
    def refresh_current_page(self):
        """Reloads the currently displayed page."""
        current_page_name = "CarbonDashboardPage" # Default
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            current_page_name = self.sidebar.get_current_page_name() or "CarbonDashboardPage"

        logging.debug(f"Refreshing page: {current_page_name}")
        # Showing the page handles destroying the old and creating the new
        self._show_page(current_page_name)

    # IVO-ONLY
    def on_closing(self):
        """Handles window close event (WM_DELETE_WINDOW or explicit close)."""
        logging.info(f"ECOHUBApp closing sequence initiated for user {self.current_user_id}...")

        # --- Pre-destroy actions ---
        # Save data ONLY if a full exit wasn't requested and we have a user ID.
        if not self._full_exit_requested and self.current_user_id:
            logging.info(f"Saving user data for {self.current_user_id} before closing.")
            save_user_data(self.current_user_id) # Handle potential errors internally
        elif self._full_exit_requested:
             logging.info("Full exit requested, skipping final data save.")
        else:
             logging.warning("Skipping data save on closing (no user ID).")

        # Stop sidebar timer safely
        try:
            if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
                self.sidebar.stop_timer()
            else:
                logging.debug("Sidebar timer stop skipped (sidebar invalid or destroyed).")
        except Exception as sidebar_e:
             logging.warning(f"Error stopping sidebar timer during on_closing: {sidebar_e}")

        # --- Initiate destruction ---
        # Call destroy() to close the Tkinter window and end mainloop
        logging.info("Initiating self.destroy().")
        try:
            self.destroy()
        except Exception as destroy_e:
             logging.error(f"Error during self.destroy(): {destroy_e}")

    # EXPENSEWISE
    def perform_full_exit(self):
        """Sets flag for full exit and starts the closing process."""
        logging.info(f"Full application exit requested by user {self.current_user_id}.")
        self._full_exit_requested = True
        # Trigger the standard closing procedure, which will handle destroy()
        self.on_closing()

# --- Sidebar Class ---
class Sidebar(tk.Frame):
    """Navigation sidebar for the main application."""

    # EXPENSEWISE
    def __init__(self, parent, show_page_callback):
        super().__init__(parent, bg=theme_colors[SIDEBAR], width=200)
        self.app = parent # Reference to the main ECOHUBApp instance
        self.show_page = show_page_callback
        self.buttons = {} # Stores {page_name: button_widget}
        self.current_page_name = None # Tracks the highlighted page
        self._timer_id = None # For the datetime update schedule

        self._build_sidebar_ui()
        self.update_datetime() # Start the clock update

    # EXPENSEWISE
    def _build_sidebar_ui(self):
        """Creates the UI elements within the sidebar."""
        # Define Structure (Dashboard, Categories, Other Pages)
        sidebar_items = [
            {"name": "CarbonDashboardPage", "text": "📊 Dashboard", "type": "page"},
            {"type": "separator"},
            {"name": "CategoriesHeader", "text": "CATEGORIES", "type": "header"},
        ]
        # Add base categories dynamically
        for cat_key, cat_info in BASE_CATEGORIES.items():
            # Determine the page name used for navigation/mapping
            page_nav_name = f"{cat_key.title().replace(' ','')}Page"
            if cat_key == "shopping": page_nav_name = "ShoppingPage" # Specific mapping for GoodsWastePage

            sidebar_items.append({
                 "name": page_nav_name,
                 "text": f"{cat_info.get('icon','')} {cat_info.get('name','Unknown')}",
                 "type": "page"
            })
        # Add remaining standard pages
        sidebar_items.extend([
            {"type": "separator"},
            {"name": "User History", "text": "📜 History", "type": "page"},
            {"name": "Settings", "text": "⚙️ Settings", "type": "page"},
        ])

        # --- Create UI Elements ---
        self.title_label = ttk.Label(self, text="ECOHUB", font=FONT_XLARGE, style="Sidebar.TLabel", anchor='w', padding=(15, 15))
        self.title_label.pack(pady=(10, 0), fill="x")

        self.datetime_label = ttk.Label(self, text="", style="DateTime.TLabel", anchor='w', padding=(15, 0))
        self.datetime_label.pack(pady=(0, 15), fill="x")

        # Create elements from config
        for item in sidebar_items:
            item_type = item["type"]
            if item_type == "separator":
                ttk.Separator(self, orient='horizontal').pack(fill='x', padx=10, pady=5)
            elif item_type == "header":
                 # Use a specific style or configure directly if needed
                 ttk.Label(self, text=item["text"], font=FONT_SMALL, style="Sidebar.TLabel", anchor='w', padding=(15, 5)).pack(fill="x", pady=(5, 0))
            elif item_type == "page":
                page_name = item["name"]
                btn = create_stylish_button(self, item["text"],
                                            # Use lambda to pass the correct page name
                                            lambda name=page_name: self.show_page(name),
                                            style="Sidebar.TButton")
                btn.pack(fill="x")
                self.buttons[page_name] = btn # Store reference

        # Spacer to push elements up
        tk.Frame(self, bg=theme_colors[SIDEBAR]).pack(expand=True, fill="y")

    # EXPENSEWISE
    def update_styles(self):
        """Updates styles of non-ttk elements or elements needing manual refresh."""
        # Called after theme switch by ECOHUBApp._configure_styles
        self.configure(bg=theme_colors[SIDEBAR])
        if hasattr(self, 'title_label') and self.title_label.winfo_exists():
             self.title_label.configure(style="Sidebar.TLabel") # Re-apply style
        if hasattr(self, 'datetime_label') and self.datetime_label.winfo_exists():
             # Recalculate color based on new theme
             date_time_fg = theme_colors[DISABLED] if self.app.current_theme == 'eco_dark' else theme_colors[ACCENT_DARKER]
             self.datetime_label.configure(style="DateTime.TLabel", foreground=date_time_fg) # Re-apply style and color

    # EXPENSEWISE
    def update_datetime(self):
        """Updates the date and time label periodically."""
        now = datetime.datetime.now()
        datetime_str = now.strftime("%a, %d %b %Y | %H:%M:%S")
        try:
            if self.datetime_label and self.datetime_label.winfo_exists():
                self.datetime_label.configure(text=datetime_str)
                # Schedule the next update using self.after
                self._timer_id = self.after(1000, self.update_datetime)
            else:
                # Label was destroyed (e.g., during theme switch before recreation)
                self._timer_id = None
                logging.debug("Sidebar datetime_label destroyed, stopping timer.")
        except tk.TclError as e:
            logging.warning(f"TclError updating sidebar time (widget likely destroyed): {e}")
            self._timer_id = None
        except Exception as e:
            logging.exception(f"Error in sidebar update_datetime: {e}")
            self._timer_id = None # Stop timer on error

    # EXPENSEWISE
    def stop_timer(self):
        """Safely cancels the pending datetime update."""
        if self._timer_id:
            try:
                self.after_cancel(self._timer_id)
                logging.debug("Sidebar datetime timer cancelled.")
            except tk.TclError: pass # Ignore error if timer ID is invalid (already cancelled/run)
            except Exception as e: logging.exception(f"Error cancelling sidebar timer: {e}")
            finally: self._timer_id = None

    # EXPENSEWISE
    def highlight_button(self, page_name):
        """Sets the 'selected' state for the corresponding sidebar button."""
        if not page_name: return
        self.current_page_name = page_name
        logging.debug(f"Highlighting sidebar button for: {page_name}")
        for name, button in self.buttons.items():
            if button and button.winfo_exists():
                try:
                    # Use 'selected' state for the active button, '!selected' for others
                    button.state(['selected'] if name == page_name else ['!selected'])
                except tk.TclError as e:
                    logging.warning(f"TclError setting state for sidebar button '{name}': {e}")

    # EXPENSEWISE
    def get_current_page_name(self):
        """Returns the name of the currently highlighted page."""
        return self.current_page_name

    # EXPENSEWISE
    def destroy(self):
        """Overrides destroy to stop the timer."""
        logging.debug("Destroying Sidebar, stopping timer.")
        self.stop_timer()
        super().destroy()

    # Define __del__ for good measure, although destroy() should handle it
    def __del__(self):
        self.stop_timer()

# --- Base Page Class (with scrollable frame setup) ---
class BasePage(tk.Frame):
    """Base class for main content pages with optional scrollable frame."""

    # IVO+GPT
    def __init__(self, parent, app):
        super().__init__(parent, bg=theme_colors[BG])
        self.app = app # Reference to the main ECOHUBApp instance
        self.app_data = app_state # Direct access to the shared data dictionary (using renamed var)
        self._mousewheel_bound_widgets = set() # Track widgets with mousewheel bound

    # IVO+GPT
    def _setup_scrollable_frame(self):
        """Creates the canvas, scrollbar, and scrollable content frame."""
        # Main canvas for scrolling
        self.canvas = tk.Canvas(self, bg=theme_colors[BG], highlightthickness=0, yscrollincrement=1)
        # Scrollbar linked to the canvas
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        # Frame inside the canvas to hold the actual page content
        self.scrollable_frame = tk.Frame(self.canvas, bg=theme_colors[BG])

        # Link canvas scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        # Place the scrollable frame within the canvas window
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", tags="scrollable_frame")

        # Grid layout for canvas and scrollbar within the BasePage frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # Bind configure events to update scroll region and frame width
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure, add='+')
        self.canvas.bind("<Configure>", self._on_canvas_configure, add='+')

        # Bind mousewheel scrolling - apply to canvas and the inner frame
        self._bind_mousewheel_to_widget(self.canvas)
        self._bind_mousewheel_to_widget(self.scrollable_frame)

        return self.scrollable_frame # Return the frame for adding content

    # IVO+GPT
    def _on_frame_configure(self, event=None):
        """Updates canvas scrollregion when the scrollable_frame resizes."""
        if hasattr(self, 'canvas') and self.canvas.winfo_exists():
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    # IVO+GPT
    def _on_canvas_configure(self, event=None):
        """Updates scrollable_frame width when the canvas resizes."""
        if hasattr(self, 'canvas') and self.canvas.winfo_exists() and hasattr(self, 'canvas_window'):
            canvas_width = self.canvas.winfo_width()
            # Ensure the item exists before configuring
            if self.canvas.find_withtag(self.canvas_window):
                 self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    # IVO+GPT
    def _bind_mousewheel_to_widget(self, widget):
        """Binds mousewheel events to a specific widget for scrolling the canvas."""
        if widget and widget.winfo_exists() and widget not in self._mousewheel_bound_widgets:
            # Ensure canvas exists before binding
            if hasattr(self, 'canvas') and self.canvas and self.canvas.winfo_exists():
                widget.bind("<MouseWheel>", self._on_mousewheel, add='+') # Windows/macOS
                widget.bind("<Button-4>", self._on_mousewheel, add='+')   # Linux scroll up
                widget.bind("<Button-5>", self._on_mousewheel, add='+')   # Linux scroll down
                self._mousewheel_bound_widgets.add(widget)
                # logging.debug(f"Mousewheel bound to {widget} for page {type(self).__name__}")

    # IVO+GPT
    def _unbind_all_mousewheel(self):
        """Unbinds mousewheel events from all tracked widgets."""
        # Iterate over a copy of the set to allow modification during iteration
        widgets_to_unbind = list(self._mousewheel_bound_widgets)
        for widget in widgets_to_unbind:
             if widget and widget.winfo_exists():
                 try:
                     widget.unbind("<MouseWheel>")
                     widget.unbind("<Button-4>")
                     widget.unbind("<Button-5>")
                     self._mousewheel_bound_widgets.discard(widget)
                     # logging.debug(f"Mousewheel unbound from {widget}")
                 except tk.TclError: pass # Ignore errors if widget destroyed concurrently
                 except Exception as e: logging.error(f"Error unbinding mousewheel from {widget}: {e}")
        self._mousewheel_bound_widgets.clear() # Ensure set is empty

    # IVO+GPT
    def _on_mousewheel(self, event):
        """Handles mousewheel events and scrolls the canvas."""
        if not hasattr(self, 'canvas') or not self.canvas or not self.canvas.winfo_exists():
            return # Canvas doesn't exist

        delta = 0
        scroll_multiplier = 3 # Units to scroll per wheel notch
        if event.num == 4: delta = -1 * scroll_multiplier # Linux up
        elif event.num == 5: delta = 1 * scroll_multiplier # Linux down
        elif hasattr(event, 'delta'): # Windows/macOS
            # Normalize delta direction
            delta = -1 * scroll_multiplier if event.delta > 0 else (1 * scroll_multiplier if event.delta < 0 else 0)

        if delta != 0:
            try:
                # Check current scroll position to avoid scrolling past limits
                scroll_info = self.canvas.yview()
                can_scroll_up = delta < 0 and scroll_info[0] > 0.0
                can_scroll_down = delta > 0 and scroll_info[1] < 1.0

                if can_scroll_up or can_scroll_down:
                    self.canvas.yview_scroll(delta, "units")
            except tk.TclError as e:
                logging.warning(f"TclError during canvas scroll: {e}") # May happen if widget destroyed during scroll

    # IVO-ONLY
    def destroy(self):
        """Overrides destroy to ensure mousewheel events are unbound."""
        logging.debug(f"Destroying {type(self).__name__}, unbinding mousewheel events.")
        self._unbind_all_mousewheel()
        super().destroy()

    # IVO-ONLY
    # Static method for formatting details
    @staticmethod
    def format_activity_details(details):
        """Creates a concise, human-readable string from activity details dict."""
        if not isinstance(details, dict):
            return str(details)

        parts = []
        # Define keys to explicitly exclude from the summary string
        exclude_keys = {k for k in details if k.startswith(('res_', 'trans_', 'food_', 'goods_', 'waste_', 'serv_', 'digital_'))}
        exclude_keys.update({k for k in details if '_period' in k or 'region' in k or 'area_type' in k})
        exclude_keys.update({"mode", "car_fuel_type", "rideshare_fuel_type", "flight_type", "flight_cabin",
                             "heat_fuel_type", "water_heater_type", "renew_type", "waste_disposal",
                             "diet_type", "area_type_retail", "local_sourcing", "packaging_level",
                             "streaming_quality", "gaming_type", "region_grid"}) # Exclude keys where value is more descriptive


        for key, value in sorted(details.items()):
            # Skip excluded keys and values that don't add info (None, empty, False)
            if key in exclude_keys or value in [None, '', False, 'N/A', 'None']:
                continue

            # Format key: Make readable
            readable_key = key.replace('_var', '').replace('_', ' ').replace(' spending', '').replace(' usage', '').replace(' type', '').replace(' gen', '').replace(' km', ' km').replace(' pkm', ' pkm').replace(' gb', ' GB').replace(' usd', '').replace(' php', ' (PHP)').replace(' hours', ' hrs').replace(' kwh', ' kWh').replace(' kg', ' kg').replace(' m2', ' m²').title().strip()

            # Format value: Handle numbers, True boolean
            formatted_value = str(value)
            if isinstance(value, bool) and value is True:
                 parts.append(readable_key) # Just show the flag name if True
                 continue
            elif isinstance(value, (int, float)):
                 try: # Format numbers nicely
                      formatted_value = f"{value:,.2f}" if isinstance(value, float) and abs(value % 1) > 1e-9 else f"{value:,}"
                 except (ValueError, TypeError): pass # Fallback to string if formatting fails

            parts.append(f"{readable_key}: {formatted_value}")

        # Join parts, provide default if no useful details found
        return " | ".join(parts) if parts else "General Entry"

# --- Placeholder Page Class ---
class PlaceholderPage(BasePage):
    """Simple page displayed for non-existent or error pages."""

    # EXPENSEWISE
    def __init__(self, parent, page_name, app):
        super().__init__(parent, app)
        # No scrolling needed here
        label = ttk.Label(self, text=f"{page_name}\n(Placeholder or Error)", style="Title.TLabel", justify=tk.CENTER, anchor=tk.CENTER)
        label.pack(expand=True, fill="both", padx=20, pady=20)

# --- CarbonDashboardPage ---
class CarbonDashboardPage(BasePage):
    """Main dashboard showing summaries and recent activity history."""

    # IVO+GPT
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.max_summary_cards_per_row = 3 # Configurable layout

        # Setup the scrollable area
        content_frame = self._setup_scrollable_frame()
        content_frame.grid_columnconfigure(0, weight=1) # Content expands horizontally
        content_frame.grid_rowconfigure(3, weight=1) # History section expands vertically

        # --- Page Content (inside content_frame) ---
        user_name = self.app_data['user_profiles'].get(self.app.current_user_id, {}).get('name', 'User')
        ttk.Label(content_frame, text=f"Welcome, {user_name}!", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(10, 5), padx=10)
        ttk.Label(content_frame, text="Your carbon footprint overview and activity history.", style="Desc.TLabel", background=theme_colors[BG]).grid(row=1, column=0, sticky="w", pady=(0, 15), padx=10)

        # Summary Cards Section
        self.create_carbon_summary_section(content_frame, row=2)

        # Full Activity History Section (Treeview)
        self.create_activity_history_section(content_frame, row=3)

        # Ensure layout is calculated for initial scroll region
        self.update_idletasks()
        self._on_frame_configure()

    # IVO+GPT
    def create_carbon_summary_section(self, parent_frame, row):
        """Creates the summary cards for each category."""
        summary_outer_frame = tk.Frame(parent_frame, bg=theme_colors[BG])
        summary_outer_frame.grid(row=row, column=0, sticky="ew", pady=(10, 15), padx=10)
        summary_outer_frame.grid_columnconfigure(0, weight=1)

        ttk.Label(summary_outer_frame, text="Category Summary", style="CardTitle.TLabel", background=theme_colors[BG]).pack(anchor="w", pady=(0, 10))

        # Grid frame for the cards
        summary_grid_frame = tk.Frame(summary_outer_frame, bg=theme_colors[BG])
        summary_grid_frame.pack(fill="x")

        # Calculate category totals
        all_activities = self.app_data.get("activities", [])
        conversion_unit = self.app_data.get("settings", {}).get("conversion", "CO2e")
        category_totals = {cat_key: 0.0 for cat_key in BASE_CATEGORIES}

        if all_activities:
            for activity in all_activities:
                cat = activity.get("category")
                fp_raw = activity.get("carbon_footprint")
                if cat in category_totals and fp_raw is not None:
                    try: category_totals[cat] += float(fp_raw)
                    except (ValueError, TypeError): pass # Ignore invalid values
        else:
             logging.info("No activities for summary.")

        # Configure grid columns based on number of categories
        num_categories = len(BASE_CATEGORIES)
        cols = min(self.max_summary_cards_per_row, num_categories) if num_categories > 0 else 1
        for i in range(cols):
            summary_grid_frame.grid_columnconfigure(i, weight=1, uniform="summary_col")

        # Create cards
        grid_row_card, grid_col_card = 0, 0
        sorted_categories = sorted(BASE_CATEGORIES.items(), key=lambda item: item[1]['name'])

        if not sorted_categories:
             ttk.Label(summary_grid_frame, text="No categories defined.", style="Card.TLabel").grid(row=0, column=0)
             return

        for category_key, category_info in sorted_categories:
            emission_value = category_totals.get(category_key, 0.0)
            card = create_card_frame(summary_grid_frame)
            card.grid(row=grid_row_card, column=grid_col_card, sticky="nsew", padx=5, pady=5)
            card.grid_columnconfigure(0, weight=1) # Allow labels inside card to align

            ttk.Label(card, text=f"{category_info['icon']} {category_info['name']}", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
            emission_label = ttk.Label(card, text=format_carbon_emission(emission_value, conversion_unit), style="Card.TLabel", font=FONT_LARGE)
            emission_label.grid(row=1, column=0, sticky="w", padx=10, pady=(5, 10))

            grid_col_card += 1
            if grid_col_card >= cols:
                grid_col_card = 0
                grid_row_card += 1

    # IVO+GPT
    def create_activity_history_section(self, parent_frame, row):
        """Creates the Treeview displaying all recorded activities."""
        history_outer_frame = tk.Frame(parent_frame, bg=theme_colors[BG])
        parent_frame.grid_rowconfigure(row, weight=1) # Allow this row to expand vertically
        history_outer_frame.grid(row=row, column=0, sticky="nsew", pady=(10, 10), padx=10)
        history_outer_frame.grid_columnconfigure(0, weight=1)
        history_outer_frame.grid_rowconfigure(1, weight=1) # Tree container expands vertically

        ttk.Label(history_outer_frame, text="Full Activity History", style="CardTitle.TLabel", background=theme_colors[BG]).grid(row=0, column=0, sticky="w", pady=(0, 5))

        # Container for Treeview + Scrollbar
        tree_container = tk.Frame(history_outer_frame, bg=theme_colors[CARD])
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        # Treeview Setup
        columns = ("timestamp", "category", "details", "footprint_kg")
        tree = ttk.Treeview(tree_container, columns=columns, show="headings", style="Treeview", height=10) # Initial height, expands
        tree_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview, style="Vertical.TScrollbar")
        tree.configure(yscrollcommand=tree_scrollbar.set)

        # Configure tags for alternating row colors ON THE TREEVIEW INSTANCE
        tree.tag_configure('oddrow', background=theme_colors[TV_ODD], foreground=theme_colors[FG])
        tree.tag_configure('evenrow', background=theme_colors[TV_EVEN], foreground=theme_colors[FG])

        # Headings and Columns
        tree.heading("timestamp", text="Date & Time", anchor=tk.W)
        tree.heading("category", text="Category", anchor=tk.W)
        tree.heading("details", text="Activity Details", anchor=tk.W)
        tree.heading("footprint_kg", text="Footprint (kg CO₂e)", anchor=tk.E) # Always show kg CO2e here

        tree.column("timestamp", width=160, anchor=tk.W, stretch=tk.NO) # Wider timestamp
        tree.column("category", width=130, anchor=tk.W, stretch=tk.NO) # Wider category
        tree.column("details", width=450, anchor=tk.W, stretch=tk.YES) # Details takes remaining space
        tree.column("footprint_kg", width=160, anchor=tk.E, stretch=tk.NO) # Wider footprint

        # Grid Treeview and Scrollbar
        tree.grid(row=0, column=0, sticky="nsew")
        tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.history_tree = tree # Store reference if needed later

        # --- Populate Treeview ---
        all_activities = self.app_data.get("activities", [])
        logging.debug(f"Populating dashboard history with {len(all_activities)} activities.")

        # Clear existing items safely
        try:
             for item in tree.get_children(): tree.delete(item)
        except tk.TclError: pass # Ignore if tree already gone

        if not all_activities:
            # Insert message directly without tags
            tree.insert("", tk.END, values=("", "No activities recorded yet.", "", ""))
            return

        # Insert data, newest first, apply tags
        try:
            # Use enumerate starting from 0 for modulo check
            for i, activity in enumerate(reversed(all_activities)):
                ts = activity.get("timestamp", "N/A")
                cat_key = activity.get("category", "unknown")
                cat_info = BASE_CATEGORIES.get(cat_key, {"icon": "?", "name": cat_key.title()})
                cat_display = f"{cat_info['icon']} {cat_info['name']}"

                details_dict = activity.get("activity_details", {})
                # Use the static BasePage method for formatting
                details_str = BasePage.format_activity_details(details_dict)

                fp_raw = activity.get("carbon_footprint")
                fp_formatted_kg = f"{float(fp_raw):,.2f}" if fp_raw is not None else "N/A"

                # Determine tag based on index (even/odd)
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                # Insert item and immediately apply the tag
                iid = tree.insert("", tk.END, values=(ts, cat_display, details_str, fp_formatted_kg), tags=(tag,))

        except Exception as e:
            logging.exception("Error populating dashboard history treeview")
            tree.insert("", tk.END, values=("Error", "Could not load history", str(e), ""))

# --- Base Class for Category Pages ---
class BaseCategoryPage(BasePage):
    """Base class for pages dedicated to a specific emission category."""

    # IVO+GPT
    def __init__(self, parent, app, category_key):
        super().__init__(parent, app)
        self.category_key = category_key
        self.category_info = BASE_CATEGORIES.get(category_key, {"name": category_key.title(), "icon": "?", "type": "category"})

        # Setup scrollable area
        content_frame = self._setup_scrollable_frame()
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_rowconfigure(3, weight=1) # History expands vertically

        # --- Page Content (within content_frame) ---
        ttk.Label(content_frame, text=f"{self.category_info['icon']} {self.category_info['name']} Footprint", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(10, 15), padx=10)

        # Analytics Summary Card
        self.analytics_frame = create_card_frame(content_frame)
        self.analytics_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15), padx=10)
        self.analytics_frame.grid_columnconfigure(0, weight=1) # Allow labels to fill width

        ttk.Label(self.analytics_frame, text="Category Summary", style="CardTitle.TLabel").grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        self.total_fp_label = ttk.Label(self.analytics_frame, text="Total: Calculating...", style="Card.TLabel", font=FONT_LARGE)
        self.total_fp_label.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="w")
        self.avg_fp_label = ttk.Label(self.analytics_frame, text="Average per Entry: Calculating...", style="Card.TLabel")
        self.avg_fp_label.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")

        # History Section
        ttk.Label(content_frame, text="Activity History", style="CardTitle.TLabel", background=theme_colors[BG]).grid(row=2, column=0, sticky="w", pady=(10, 5), padx=10)

        # History Treeview Container
        tree_container = tk.Frame(content_frame, bg=theme_colors[CARD])
        tree_container.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        # History Treeview
        columns = ("timestamp", "details", "footprint")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", style="Treeview", height=8)
        self.tree_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)

        # Configure row tags on the treeview instance
        self.tree.tag_configure('oddrow', background=theme_colors[TV_ODD], foreground=theme_colors[FG])
        self.tree.tag_configure('evenrow', background=theme_colors[TV_EVEN], foreground=theme_colors[FG])

        # Headings and Columns
        self.tree.heading("timestamp", text="Date & Time", anchor=tk.W)
        self.tree.heading("details", text="Activity Details", anchor=tk.W)
        # Footprint heading updated dynamically in refresh_data based on settings
        self.tree.heading("footprint", text="Footprint", anchor=tk.E)

        self.tree.column("timestamp", width=160, anchor=tk.W, stretch=tk.NO)
        self.tree.column("details", width=400, anchor=tk.W, stretch=tk.YES) # Details takes space
        self.tree.column("footprint", width=160, anchor=tk.E, stretch=tk.NO)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scrollbar.grid(row=0, column=1, sticky="ns")

        # --- Initial Data Load ---
        self.refresh_data()
        self.update_idletasks()
        self._on_frame_configure()

    # IVO-ONLY
    def get_category_activities(self):
        """Filters all activities for the current category."""
        all_activities = self.app_data.get("activities", [])
        # List comprehension for concise filtering
        return [act for act in all_activities if isinstance(act, dict) and act.get("category") == self.category_key]

    # IVO+GPT
    def load_activity_history(self):
        """Populates the Treeview with filtered history."""

        if not hasattr(self, 'tree') or not self.tree.winfo_exists(): return

        # Clear existing items safely
        try:
             for item in self.tree.get_children(): self.tree.delete(item)
        except tk.TclError: pass

        category_activities = self.get_category_activities()
        conversion_unit = self.app_data.get("settings", {}).get("conversion", "CO2e")
        # Update tree heading based on current unit
        unit_suffix = "kg CO₂e"
        if "Trees" in conversion_unit: unit_suffix = "Trees/yr"
        elif "Cars" in conversion_unit: unit_suffix = "Cars/yr"
        self.tree.heading("footprint", text=f"Footprint ({unit_suffix})", anchor=tk.E)

        logging.debug(f"{self.category_key}Page: Loading {len(category_activities)} history items. Unit: {conversion_unit}")

        if not category_activities:
             self.tree.insert("", tk.END, values=("", f"No {self.category_key} activities recorded yet.", ""))
             return

        # IVO-ONLY
        # Insert data, newest first, applying tags
        try:
            for i, activity in enumerate(reversed(category_activities)):
                ts = activity.get("timestamp", "N/A")
                details_dict = activity.get("activity_details", {})
                details_str = BasePage.format_activity_details(details_dict) # Use static formatter
                fp_raw = activity.get("carbon_footprint")
                fp_formatted = format_carbon_emission(fp_raw, conversion_unit) if fp_raw is not None else "N/A"

                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                # Insert with tag
                iid = self.tree.insert("", tk.END, values=(ts, details_str, fp_formatted), tags=(tag,))

        except Exception as e:
            logging.exception(f"Error populating history for {self.category_key}")
            try: # Clear tree before inserting error message
                for item in self.tree.get_children(): self.tree.delete(item)
            except tk.TclError: pass
            self.tree.insert("", tk.END, values=("Error", "Could not load history", str(e)))

    # IVO-ONLY
    def update_analytics(self):
        """Calculates and displays summary stats for the category."""
        # Check if widgets exist before updating
        required_widgets = ['analytics_frame', 'total_fp_label', 'avg_fp_label']
        if not all(hasattr(self, attr) and getattr(self, attr) and getattr(self, attr).winfo_exists() for attr in required_widgets):
            logging.warning(f"{self.category_key}Page: Analytics widgets missing, cannot update.")
            return

        try:
            category_activities = self.get_category_activities()
            activity_count = len(category_activities)

            # Calculate total and average footprint in kg CO2e
            total_co2e = sum(float(act.get("carbon_footprint", 0) or 0) for act in category_activities)
            avg_co2e = (total_co2e / activity_count) if activity_count > 0 else 0.0

            # Get user's preferred display unit
            conversion_unit = self.app_data.get("settings", {}).get("conversion", "CO2e")

            # Format values using the selected unit
            formatted_total = format_carbon_emission(total_co2e, conversion_unit)
            formatted_avg = format_carbon_emission(avg_co2e, conversion_unit)

            # Update labels
            self.total_fp_label.configure(text=f"Total: {formatted_total}")
            self.avg_fp_label.configure(text=f"Average / Entry: {formatted_avg} ({activity_count} entries)")

            logging.debug(f"{self.category_key}Page: Analytics updated - Total={formatted_total}, Avg={formatted_avg}")

        except Exception as e:
            logging.exception(f"Error updating analytics for {self.category_key}")
            # Update labels to show error state
            if hasattr(self, 'total_fp_label') and self.total_fp_label.winfo_exists():
                self.total_fp_label.configure(text="Total: Error")
            if hasattr(self, 'avg_fp_label') and self.avg_fp_label.winfo_exists():
                 self.avg_fp_label.configure(text="Average: Error")

    # IVO+GPT
    def refresh_data(self):
        """Reloads and recalculates data for the page."""
        logging.debug(f"Refreshing data for {self.category_key} page.")
        # Update analytics based on current data
        self.update_analytics()
        # Reload history using current conversion setting
        self.load_activity_history()

        # Update layout and scroll region
        self.update_idletasks()
        if hasattr(self, 'canvas') and self.canvas.winfo_exists():
             self._on_frame_configure()

# --- Specific Category Page Classes (Inherit from Base) ---
class ResidentialPage(BaseCategoryPage):

    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "residential")

class TravelPage(BaseCategoryPage):

    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "travel")

class FoodPage(BaseCategoryPage):

    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "food")

class GoodsWastePage(BaseCategoryPage):
    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "shopping")

class ServicesPage(BaseCategoryPage):
    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "services")

class DigitalPage(BaseCategoryPage):
    # IVO-ONLY
    def __init__(self, parent, app): super().__init__(parent, app, "digital")

# IVO+GPT
# --- UserHistoryPage Class ---
class UserHistoryPage(BasePage):
    """Displays the log of user actions (not carbon activities)."""
    def __init__(self, parent, app):
        super().__init__(parent, app)
        # This page likely doesn't need scrolling unless MAX_ACTIVITY_LOG_SIZE is huge
        self.grid_rowconfigure(1, weight=1) # Treeview container expands
        self.grid_columnconfigure(0, weight=1)

        ttk.Label(self, text="User Action History", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(10, 15), padx=10)

        # Treeview Container
        tree_container = tk.Frame(self, bg=theme_colors[CARD])
        tree_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        # Treeview Setup
        columns = ("timestamp", "action")
        tree = ttk.Treeview(tree_container, columns=columns, show="headings", style="Treeview", height=15) # Fixed height usually fine
        tree.heading("timestamp", text="Timestamp", anchor=tk.W)
        tree.heading("action", text="Action Description", anchor=tk.W)

        tree.column("timestamp", width=170, anchor=tk.W, stretch=tk.NO) # Slightly wider
        tree.column("action", width=600, anchor=tk.W, stretch=tk.YES) # Action takes space

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview, style="Vertical.TScrollbar")
        tree.configure(yscrollcommand=scrollbar.set)

        # Configure row tags on the instance
        tree.tag_configure('oddrow', background=theme_colors[TV_ODD], foreground=theme_colors[FG])
        tree.tag_configure('evenrow', background=theme_colors[TV_EVEN], foreground=theme_colors[FG])

        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # --- Populate Treeview ---
        activity_log = self.app_data.get("activity_log", [])
        if not isinstance(activity_log, list): # Ensure it's a list
            logging.error("User activity log data is not a list, resetting.")
            activity_log = []
            self.app_data["activity_log"] = activity_log # Fix in memory

        logging.debug(f"Populating user history page with {len(activity_log)} entries.")

        # Clear existing items safely
        try:
             for item in tree.get_children(): tree.delete(item)
        except tk.TclError: pass

        if not activity_log:
            tree.insert("", tk.END, values=("", "No user actions logged yet."))
        else:
            try:
                # Insert data, newest first, apply tags
                for i, log_entry in enumerate(reversed(activity_log)):
                    tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                    if isinstance(log_entry, dict):
                        ts = log_entry.get('timestamp', 'N/A')
                        action = log_entry.get('action', 'N/A')
                        tree.insert("", tk.END, values=(ts, action), tags=(tag,))
                    else: # Handle unexpected format
                        tree.insert("", tk.END, values=("Invalid Entry", str(log_entry)), tags=(tag,))
            except Exception as e:
                logging.exception("Error populating user action history log treeview")
                try: # Clear before inserting error
                    for item in tree.get_children(): tree.delete(item)
                except tk.TclError: pass
                tree.insert("", tk.END, values=("Error", f"Could not load log: {e}"))

# --- SettingsPage Class ---
class SettingsPage(BasePage):
    """Allows user to change theme, units, and manage profile/data."""

    # EXPENSEWISE
    def __init__(self, parent, app):
        super().__init__(parent, app)
        # Settings page might benefit from scrolling if more options added
        # content_frame = self._setup_scrollable_frame()
        # If not scrollable, grid directly into self:
        content_frame = self
        content_frame.grid_columnconfigure(0, weight=1) # Allow cards to expand width

        self._create_ui_elements(content_frame)

    # EXPENSEWISE
    def _create_ui_elements(self, parent_frame):
        """Creates the settings widgets within the specified parent."""
        # Clear previous elements if recreating UI
        for widget in parent_frame.winfo_children():
             # Don't destroy canvas/scrollbar if using scrollable frame
             if widget not in getattr(self, '_scroll_components', []):
                 widget.destroy()

        current_row = 0
        ttk.Label(parent_frame, text="Settings", style="Title.TLabel").grid(row=current_row, column=0, sticky="w", pady=(10, 20), padx=10); current_row += 1 # Added padding

        # --- Theme Settings Card ---
        theme_frame = create_card_frame(parent_frame)
        theme_frame.grid(row=current_row, column=0, sticky="ew", padx=10, pady=10); current_row += 1
        theme_frame.grid_columnconfigure(1, weight=1) # Allow radio button text area to expand if needed
        ttk.Label(theme_frame, text="Appearance", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))
        # Use theme setting from app_state
        self.theme_var = tk.StringVar(value=self.app_data.get("settings", {}).get("theme", "eco_dark"))
        light_rb = ttk.Radiobutton(theme_frame, text="Light Theme", variable=self.theme_var, value="eco_light", command=self._change_theme, style="Card.TRadiobutton")
        light_rb.grid(row=1, column=0, sticky="w", padx=(20, 10), pady=2)
        dark_rb = ttk.Radiobutton(theme_frame, text="Dark Theme", variable=self.theme_var, value="eco_dark", command=self._change_theme, style="Card.TRadiobutton")
        dark_rb.grid(row=2, column=0, sticky="w", padx=(20, 10), pady=(2, 10))

        # --- Conversion Settings Card ---
        conv_frame = create_card_frame(parent_frame)
        conv_frame.grid(row=current_row, column=0, sticky="ew", padx=10, pady=10); current_row += 1
        conv_frame.grid_columnconfigure(0, weight=1) # Allow combobox/label to expand
        ttk.Label(conv_frame, text="Footprint Display Unit", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.conversion_var = tk.StringVar(value=self.app_data.get("settings", {}).get("conversion", "CO2e"))
        conversion_options = ["CO2e", "Trees (Absorbed CO2 per Year)", "Cars (Emitted CO2 per Year)"]
        # Use themed Combobox style
        conversion_combo = ttk.Combobox(conv_frame, textvariable=self.conversion_var, values=conversion_options, state='readonly', style='TCombobox', width=35) # Adjust width
        conversion_combo.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 5))
        conversion_combo.bind("<<ComboboxSelected>>", self._change_conversion_unit)
        # Use themed Label style
        ttk.Label(conv_frame, text="Select how carbon footprint values are displayed across the app.", style="CardDesc.TLabel", wraplength=350).grid(row=2, column=0, sticky="w", padx=20, pady=(0, 10))

        # --- User Profile Actions Card ---
        action_frame = create_card_frame(parent_frame)
        action_frame.grid(row=current_row, column=0, sticky="ew", padx=10, pady=10); current_row += 1
        ttk.Label(action_frame, text="Profile & Application Actions", style="CardTitle.TLabel").pack(padx=10, pady=(10, 5), anchor='w')
        # Frame to hold the buttons, allowing them to flow or be arranged
        buttons_frame = tk.Frame(action_frame, bg=theme_colors[CARD])
        buttons_frame.pack(fill="x", padx=10, pady=(5, 10))

        # Create buttons with standard style
        switch_button = create_stylish_button(buttons_frame, "Switch Profile", self._switch_user)
        switch_button.pack(side=tk.LEFT, padx=(0, 5), pady=2)
        reset_button = create_stylish_button(buttons_frame, "Reset Data", self._reset_data)
        reset_button.pack(side=tk.LEFT, padx=5, pady=2)
        delete_button = create_stylish_button(buttons_frame, "Delete Profile", self._delete_user)
        delete_button.pack(side=tk.LEFT, padx=5, pady=2)
        exit_button = create_stylish_button(buttons_frame, "Exit Application", self._exit_application)
        exit_button.pack(side=tk.LEFT, padx=5, pady=2)

    # EXPENSEWISE
    def _change_theme(self):
        new_theme = self.theme_var.get()
        logging.info(f"Theme selection changed to: {new_theme}")
        self.app.switch_theme(new_theme)
        self._create_ui_elements(self)
        self.configure(bg=theme_colors["background"])

    # GUTIERREZ+KATSUYA
    def _change_conversion_unit(self, event=None):
        """Handles conversion unit selection change."""
        if not self.app or not self.app.winfo_exists(): return # App already closed?
        user_id = self.app.current_user_id
        if not user_id:
            logging.error("Settings: Cannot change conversion, current user ID missing.")
            return

        new_unit = self.conversion_var.get()
        valid_units = ["CO2e", "Trees (Absorbed CO2 per Year)", "Cars (Emitted CO2 per Year)"]
        if new_unit not in valid_units:
            logging.warning(f"Invalid conversion unit selected: {new_unit}")
            # Optionally reset var to previous value
            self.conversion_var.set(self.app_data["settings"].get("conversion", "CO2e"))
            return

        # Update setting in memory
        self.app_data["settings"]["conversion"] = new_unit
        # Save the setting immediately
        if save_user_data(user_id):
            logging.info(f"Conversion unit changed to: {new_unit} and saved.")
            log_activity(f"Display unit changed to {new_unit}")
            # Refresh current page to reflect the new unit ONLY if save successful
            self.app.refresh_current_page()
        else:
            # Save failed (error shown by save_user_data), revert change in memory
            # A bit complex, might need to reload settings? For now, just log.
            logging.error(f"Failed to save conversion unit '{new_unit}'. Reverting might be needed.")
            # Optionally revert the Combobox selection visually?

    # EXPENSEWISE
    def _switch_user(self):
        """Closes the current user session and returns to profile selection."""
        logging.info("Settings: Switch User action initiated.")
        if self.app and self.app.winfo_exists():
            # Trigger the app's closing sequence (saves data if needed)
            # on_closing will eventually call destroy()
            self.app.on_closing()
        else:
            logging.error("Settings: Cannot switch user, main app instance missing.")
            messagebox.showerror("Error", "Application state error. Cannot switch user.", parent=self)

    # EXPENSEWISE
    def _reset_data(self):
        """Resets activity and log data for the current user after confirmation."""
        if not self.app or not self.app.winfo_exists(): return
        user_id = self.app.current_user_id
        if not user_id:
             messagebox.showerror("Error", "Cannot determine current user.", parent=self)
             return

        if messagebox.askyesno("Reset Data Confirmation",
                               "This will permanently clear ALL recorded activities and the user action history for this profile.\n\n"
                               "Settings (like theme) and the profile itself will remain.\n\n"
                               "This cannot be undone. Proceed?",
                               icon='warning', parent=self):
            try:
                logging.warning(f"Resetting activity and log data for user {user_id}")
                # Clear data in memory
                self.app_data["activities"] = []
                self.app_data["activity_log"] = []

                # Save empty lists to files
                # Note: _save_json_data handles directory creation
                activity_file = get_user_data_file_path(user_id, "activities")
                log_file = get_user_data_file_path(user_id, "activity_log")
                save_act_ok = _save_json_data(activity_file, [])
                save_log_ok = _save_json_data(log_file, [])

                if not save_act_ok or not save_log_ok:
                    # Error message handled by _save_json_data indirectly via save_user_data
                    # Best effort: Try to reload original data? Risky. Just notify.
                    messagebox.showerror("Save Error", "Failed to save reset data files. Data might be inconsistent.", parent=self)
                    # Attempt reload to restore memory state?
                    try: load_user_data(user_id)
                    except: logging.error("Failed reload after reset save error.")
                    return # Stop

                # Log the reset action itself
                log_activity("Reset user activity and log data")
                # Save the activity log *again* to include the reset action
                _save_json_data(log_file, self.app_data["activity_log"]) # Ignore error here?

                messagebox.showinfo("Data Reset", "User activity and log data reset successfully.", parent=self)
                # Refresh the current page (e.g., dashboard) to show empty state
                self.app.refresh_current_page()

            except Exception as e:
                logging.exception("Error during data reset process")
                messagebox.showerror("Error", f"Could not complete data reset:\n{e}", parent=self)

    # EXPENSEWISE
    def _delete_user(self):
        """Deletes the current user profile and all associated data after confirmation."""
        if not self.app or not self.app.winfo_exists(): return
        user_id = self.app.current_user_id
        if not user_id:
             messagebox.showerror("Error", "Cannot determine current user.", parent=self)
             return

        user_name = self.app_data.get("user_profiles", {}).get(user_id, {}).get("name", f"ID: {user_id}")

        if messagebox.askyesno("Delete Profile Confirmation",
                               f"Delete profile '{user_name}'?\n\n"
                               "This will permanently delete the profile and ALL associated data (activities, logs, settings).\n\n"
                               "This cannot be undone. Are you absolutely sure?",
                               icon='warning', parent=self):
            try:
                logging.warning(f"Attempting deletion of user: {user_name} (ID: {user_id})")

                # 1. Remove from profiles dictionary in memory
                profiles = self.app_data.get("user_profiles", {})
                if user_id in profiles:
                    del profiles[user_id]
                else:
                     logging.warning(f"User ID {user_id} not found in profiles dictionary during delete.")

                # 2. Save updated profiles list
                if not save_user_profiles_to_csv():
                    logging.error("Failed to save updated profiles CSV after deleting user. Aborting.")
                    # Attempt to restore profile in memory
                    load_user_profiles_from_csv() # Reload from (hopefully unchanged) file
                    messagebox.showerror("Save Error", "Could not update user list file. Profile deletion aborted.", parent=self)
                    return

                # 3. Delete associated user data files
                data_files_to_delete = [get_user_data_file_path(user_id, dt) for dt in ["settings", "activities", "activity_log"]]
                deletion_errors = []
                for file_path in data_files_to_delete:
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                            logging.info(f"Deleted user data file: {file_path}")
                        except OSError as e:
                            logging.error(f"Could not delete file {file_path}: {e}")
                            deletion_errors.append(os.path.basename(file_path))

                # 4. Show result message
                if deletion_errors:
                    messagebox.showwarning("Deletion Warning",
                                           f"Profile '{user_name}' removed, but failed to delete data files:\n" +
                                           "\n".join(deletion_errors) +
                                           "\nManual cleanup may be needed.", parent=self)
                else:
                    messagebox.showinfo("User Deleted", f"Profile '{user_name}' and data deleted successfully.", parent=self)

                # 5. Close the app window to return to account selection
                # Set the flag to prevent final save on closing
                self.app._full_exit_requested = True # Treat deletion like a switch/exit
                self.app.on_closing()

            except Exception as e:
                logging.exception(f"Error deleting user {user_id}")
                messagebox.showerror("Error", f"An error occurred deleting the profile:\n{e}", parent=self)
                 # Attempt to reload profiles if something went wrong mid-process
                try: load_user_profiles_from_csv()
                except: pass

    # EXPENSEWISE
    def _exit_application(self):
        """Initiates a full application exit."""
        logging.info("Settings: Exit Application action initiated.")
        if self.app and self.app.winfo_exists():
            # Call the app's method to handle full exit cleanly
            self.app.perform_full_exit()
        else:
            logging.error("Settings: Cannot exit, main app instance missing.")
            messagebox.showerror("Error", "Application state error. Cannot exit.", parent=self)


# --- Add Carbon Footprint Activity Dialog ---
class AddCarbonFootprintActivityDialog(tk.Toplevel):
    """Modal dialog for entering carbon footprint activities."""

    # IVO+GPT
    def __init__(self, parent_app):
        super().__init__(parent_app)
        self.app = parent_app
        self.app_data = app_state # Use renamed global state dict
        self.factors = self.app_data.get("emission_factors", DEFAULT_EMISSION_FACTORS)

        # Dialog config
        self.configure(bg=theme_colors[DLG_BG])
        self.title("Add Carbon Footprint Activity")
        self.geometry("750x780") # Can adjust size as needed
        self.resizable(True, True)
        self.transient(parent_app)
        self.grab_set() # Make modal

        # Dialog-specific styles
        self.dialog_style = ttk.Style(self)
        self.dialog_style.theme_use('clam')
        self._configure_dialog_styles()

        # Store references to scrollable components for cleanup
        self._scroll_widgets_by_tab = {}
        # Store input variable dicts per category
        self.activity_vars = {cat_key: {} for cat_key in BASE_CATEGORIES}
        # Store specific widget references for dynamic UI changes
        self.travel_widgets = {}
        self.res_heat_widgets = {}
        self.res_water_widgets = {}
        # Store food/spending mappings as instance attributes
        self._food_inputs_map = {
            "beef_kg": ("food_prod_beef_kg_kg", "food_prod_lamb_kg_kg"), # Tuple for averaging
            "pork_kg": "food_prod_pork_kg_kg",
            "poultry_kg": "food_prod_poultry_kg_kg",
            "seafood_kg": "food_prod_seafood_kg_kg",
            "dairy_kg": "food_prod_dairy_kg_kg",
            "eggs_kg": "food_prod_eggs_kg_kg",
            # Combine Veg/Fruits/Grains/Legumes for simplicity if desired, or keep separate
            "veg_fruit_kg": ("food_prod_vegetables_kg_kg", "food_prod_fruits_kg_kg"),
            "grains_legumes_kg": ("food_prod_grains_kg_kg", "food_prod_legumes_kg_kg")
        }
        self._food_inputs_labels = { # For validation messages
            "beef_kg": "Beef / Lamb", "pork_kg": "Pork", "poultry_kg": "Poultry",
            "seafood_kg": "Fish & Seafood", "dairy_kg": "Dairy", "eggs_kg": "Eggs",
            "veg_fruit_kg": "Vegetables & Fruits", "grains_legumes_kg": "Grains & Legumes"
        }
        self._spending_cats_map = {
            "clothing_spending": "spending_clothing_usd",
            "electronics_spending": "spending_electronics_usd",
            "appliances_spending": "spending_appliances_usd",
            "furniture_spending": "spending_furniture_usd",
            "other_spending": "spending_other_goods_usd"
        }
        self._spending_cats_labels = { # For validation messages
             "clothing_spending": "Clothing Spending", "electronics_spending": "Electronics Spending",
             "appliances_spending": "Appliances Spending", "furniture_spending": "Furniture Spending",
             "other_spending": "Other Goods Spending"
        }

        # Main container for padding
        main_container = tk.Frame(self, bg=theme_colors[DLG_BG], padx=15, pady=15)
        main_container.pack(expand=True, fill="both")
        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_rowconfigure(0, weight=1) # Notebook area expands

        # Notebook for categories
        self.notebook = ttk.Notebook(main_container, style='TNotebook')
        self.notebook.grid(row=0, column=0, sticky="nsew", pady=(0, 15))

        # Create tabs dynamically
        self.tabs = {} # Stores {cat_key: scrollable_content_frame}
        for cat_key, cat_info in BASE_CATEGORIES.items():
            # Container frame added to notebook (holds canvas/scrollbar)
            tab_container_frame = tk.Frame(self.notebook, bg=theme_colors[DLG_BG])
            tab_container_frame.pack(expand=True, fill="both")

            # Setup scrollable components inside the container
            canvas, scrollbar, scrollable_content_frame = self._setup_scrollable_tab(tab_container_frame)

            # Store references (use frame widget itself as key for simplicity)
            self._scroll_widgets_by_tab[scrollable_content_frame] = {'canvas': canvas, 'scrollbar': scrollbar}
            self.tabs[cat_key] = scrollable_content_frame # Store content frame ref

            # Add container to notebook
            self.notebook.add(tab_container_frame, text=f" {cat_info['icon']} {cat_info['name']} ")

            # Populate the scrollable content frame with widgets
            # Map category key to the correct creation function name
            func_suffix = cat_key if cat_key != "shopping" else "goods_waste"
            create_func = getattr(self, f"create_{func_suffix}_tab_widgets", None)
            if create_func and callable(create_func):
                create_func(scrollable_content_frame) # Pass the inner frame as parent
            else:
                logging.warning(f"Widget creation function for category '{cat_key}' not found.")
                ttk.Label(scrollable_content_frame, text=f"Input form for {cat_info['name']} not implemented.", style="Dialog.TLabel").grid(row=0, column=0, pady=10, padx=10)

        # Action Buttons Frame
        action_frame = tk.Frame(main_container, bg=theme_colors[DLG_BG])
        action_frame.grid(row=1, column=0, sticky="sew", pady=(10, 0))
        action_frame.grid_columnconfigure(0, weight=1) # Push buttons right

        self.submit_button = ttk.Button(action_frame, text="Add Activity", command=self.submit_activity, style="Dialog.TButton", width=15)
        self.submit_button.grid(row=0, column=1, padx=(10,5), ipady=3)
        self.cancel_button = ttk.Button(action_frame, text="Cancel", command=self.destroy, style="Cancel.Dialog.TButton", width=10)
        self.cancel_button.grid(row=0, column=2, padx=(0,0), ipady=3)

        # Final setup
        self.update_idletasks()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change) # Bind after creation
        self._set_initial_focus() # Focus first field in active tab
        self._center_dialog(parent_app)

        self.bind("<Escape>", lambda e: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy) # Use destroy for cleanup

        self.wait_window(self) # Block until dialog is closed

    # EXPENSEWISE
    def _configure_dialog_styles(self):
        """Configures ttk styles specifically for the dialog elements."""
        # Styles moved here, using theme_colors dict keys
        bg = theme_colors[DLG_BG]; fg = theme_colors[DLG_FG]; card_bg = theme_colors[DLG_CARD]
        accent = theme_colors[ACCENT]; accent_darker = theme_colors[ACCENT_DARKER]
        disabled = theme_colors[DISABLED]; btn_fg = theme_colors[BTN_FG]
        red = theme_colors[RED]; desc_fg = theme_colors[LBL_DESC_FG]

        self.dialog_style.configure('Dialog.TLabel', background=card_bg, foreground=fg, font=FONT_NORMAL)
        self.dialog_style.configure('DialogBold.TLabel', background=card_bg, foreground=fg, font=FONT_BOLD)
        self.dialog_style.configure('DialogDesc.TLabel', background=card_bg, foreground=desc_fg, font=FONT_DESC)
        self.dialog_style.configure('Unit.TLabel', background=card_bg, foreground=desc_fg, font=FONT_SMALL)
        self.dialog_style.configure('SectionHeader.TLabel', background=card_bg, foreground=fg, font=FONT_BOLD)

        self.dialog_style.configure('Dialog.TEntry', fieldbackground=card_bg, foreground=fg, insertcolor=fg, borderwidth=1, relief=tk.SOLID, bordercolor=disabled, padding=5)
        self.dialog_style.map('Dialog.TEntry', bordercolor=[('focus', accent)])
        self.dialog_style.configure('Dialog.TCombobox', fieldbackground=card_bg, foreground=fg, selectbackground=card_bg, selectforeground=fg, arrowcolor=fg, borderwidth=1, relief=tk.SOLID, bordercolor=disabled, padding=5)
        self.dialog_style.map('Dialog.TCombobox', bordercolor=[('focus', accent)], fieldbackground=[('readonly', card_bg)])
        # Use the specific Dialog Checkbutton style configured in main app styles
        # self.dialog_style.configure('Dialog.TCheckbutton', background=card_bg, foreground=fg, font=FONT_NORMAL)
        # self.dialog_style.map('Dialog.TCheckbutton', indicatorcolor=[('selected', accent)])

        self.dialog_style.configure('Dialog.TButton', background=accent, foreground=btn_fg, font=FONT_BOLD, padding=(10, 5))
        self.dialog_style.map('Dialog.TButton', background=[('active', accent_darker)])
        self.dialog_style.configure('Cancel.Dialog.TButton', background=disabled, foreground=fg, font=FONT_BOLD, padding=(10, 5))
        self.dialog_style.map('Cancel.Dialog.TButton', background=[('active', red)], foreground=[('active', btn_fg)])

    # EXPENSEWISE
    def _on_tab_change(self):
        """Callback for when the notebook tab changes."""
        self._set_initial_focus()
        # Also ensure the new tab's scroll region is correct
        try:
            selected_widget_name = self.notebook.select()
            if not selected_widget_name: return
            container_frame = self.notebook.nametowidget(selected_widget_name)
            # Find the scrollable frame associated with this container
            for frame, info in self._scroll_widgets_by_tab.items():
                 if info['canvas'].master == container_frame:
                      self.update_idletasks() # Ensure widgets drawn
                      self._on_frame_configure(canvas=info['canvas']) # Update scroll region
                      break
        except Exception as e:
            logging.warning(f"Error handling tab change: {e}")

    # EXPENSEWISE
    def _set_initial_focus(self, event=None):
        """Sets focus to the first logical input field in the active tab."""
        try:
            selected_tab_widget_name = self.notebook.select()
            if not selected_tab_widget_name: return

            container_frame = self.notebook.nametowidget(selected_tab_widget_name)
            # Find the corresponding scrollable_content_frame
            scrollable_content_frame = None
            canvas = None
            for frame, info in self._scroll_widgets_by_tab.items():
                 if info['canvas'].master == container_frame:
                      scrollable_content_frame = frame
                      canvas = info['canvas']
                      break

            if not scrollable_content_frame or not scrollable_content_frame.winfo_exists():
                logging.warning(f"Cannot find active scrollable frame for tab '{selected_tab_widget_name}' to set focus.")
                self.notebook.focus_set() # Fallback focus
                return

            # Find first focusable Entry or Combobox
            first_input = self._find_first_focusable(scrollable_content_frame)

            if first_input:
                first_input.focus_set()
                # Scroll into view after a short delay
                if canvas: self.after(50, lambda c=canvas, w=first_input: self._scroll_widget_into_view(c, w))
                # Select text in Entry for easy replacement
                if isinstance(first_input, ttk.Entry):
                    try: first_input.select_range(0, tk.END)
                    except tk.TclError: pass # Ignore selection errors
            else:
                self.notebook.focus_set() # Fallback if no inputs found

        except Exception as e:
            logging.warning(f"Could not set initial focus in dialog: {e}", exc_info=True)
            try: self.notebook.focus_set() # Final fallback attempt
            except Exception: pass

    # IVO+GPT
    def _find_first_focusable(self, parent_widget):
         """Recursively find the first focusable Entry/Combobox."""
         for widget in parent_widget.winfo_children():
              if isinstance(widget, (ttk.Entry, ttk.Combobox)) and 'disabled' not in widget.state():
                  return widget
              elif isinstance(widget, tk.Frame): # Recurse into sub-frames
                  found = self._find_first_focusable(widget)
                  if found: return found
         return None

    # IVO+GPT
    def _scroll_widget_into_view(self, canvas, widget):
        """Scrolls the canvas so the widget is visible."""
        try:
            if not canvas or not canvas.winfo_exists() or not widget or not widget.winfo_exists(): return
            canvas.update_idletasks()

            widget_y = widget.winfo_y()
            widget_h = widget.winfo_height()
            canvas_h = canvas.winfo_height()

            # Get current scroll view (fractions)
            scroll_top_frac, scroll_bottom_frac = canvas.yview()
            # Calculate total height of scrollable content
            scroll_region_str = canvas.cget('scrollregion')
            if not scroll_region_str: return
            scroll_region = list(map(float, scroll_region_str.split()))
            if len(scroll_region) < 4: return
            content_height = scroll_region[3] - scroll_region[1]
            if content_height <= 0: return

            # Calculate current view bounds in pixels
            view_top_y = scroll_top_frac * content_height
            view_bottom_y = view_top_y + canvas_h # Approximation, could use scroll_bottom_frac

            # Determine target scroll fraction
            target_fraction = scroll_top_frac
            needs_scroll = False
            if widget_y < view_top_y: # Widget above view
                target_fraction = widget_y / content_height
                needs_scroll = True
            elif widget_y + widget_h > view_bottom_y: # Widget below view (or partially)
                 # Scroll to bring bottom of widget to bottom of view
                 target_fraction = (widget_y + widget_h - canvas_h) / content_height
                 needs_scroll = True

            if needs_scroll:
                target_fraction = max(0.0, min(target_fraction, 1.0)) # Clamp fraction
                canvas.yview_moveto(target_fraction)

        except (tk.TclError, ValueError, IndexError, AttributeError, TypeError) as e:
             logging.warning(f"Error scrolling widget {widget} into view: {e}")

    # EXPENSEWISE
    def _center_dialog(self, parent):
        """Centers the dialog relative to the parent window."""
        try:
            self.update_idletasks()
            parent_x, parent_y = parent.winfo_rootx(), parent.winfo_rooty()
            parent_w, parent_h = parent.winfo_width(), parent.winfo_height()
            dialog_w, dialog_h = self.winfo_width(), self.winfo_height()

            x = parent_x + (parent_w // 2) - (dialog_w // 2)
            y = parent_y + (parent_h // 3) - (dialog_h // 2) # Place higher than center

            screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
            margin = 15 # Margin from screen edge
            x = max(margin, min(x, screen_w - dialog_w - margin))
            y = max(margin, min(y, screen_h - dialog_h - margin))

            self.geometry(f"+{x}+{y}")
        except Exception as e:
             logging.error(f"Error centering dialog: {e}")

    # IVO+GPT
    # --- Scrolling Setup ---
    def _setup_scrollable_tab(self, parent_container):
        """Creates scrollable components for a notebook tab."""
        content_bg = theme_colors[DLG_CARD] # Use dialog card color for content area
        canvas = tk.Canvas(parent_container, bg=content_bg, highlightthickness=0, yscrollincrement=1)
        scrollbar = ttk.Scrollbar(parent_container, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
        # Inner frame for actual widgets, with padding
        scrollable_frame = tk.Frame(canvas, bg=content_bg, padx=20, pady=15)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags="scrollable_frame")

        parent_container.grid_rowconfigure(0, weight=1)
        parent_container.grid_columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure grid inside the scrollable frame (allow input column to expand)
        scrollable_frame.grid_columnconfigure(1, weight=1)

        # Bind events for scroll region and width updates
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: self._on_frame_configure(e, c), add='+')
        canvas.bind("<Configure>", lambda e, c=canvas, cw=canvas_window: self._on_canvas_configure(e, c, cw), add='+')

        # Bind mousewheel events (BasePage methods handle this)
        self._bind_mousewheel_to_widget(canvas)
        self._bind_mousewheel_to_widget(scrollable_frame)

        return canvas, scrollbar, scrollable_frame

    # EXPENSEWISE
    def _on_frame_configure(self, event, canvas):
        """Callback for scrollable_frame <Configure>."""
        if canvas and canvas.winfo_exists():
            canvas.configure(scrollregion=canvas.bbox("all"))

    # EXPENSEWISE
    def _on_canvas_configure(self, event, canvas, canvas_window):
        """Callback for Canvas <Configure>."""
        if canvas and canvas.winfo_exists() and canvas_window:
            canvas_width = canvas.winfo_width()
            # Check item exists before config
            if canvas.find_withtag(canvas_window):
                 canvas.itemconfig(canvas_window, width=canvas_width)

    # IVO+GPT
    # Inherit mousewheel binding/handling logic from BasePage? No, Dialog is Toplevel. Re-implement.
    def _bind_mousewheel_to_widget(self, widget):
        """Binds mousewheel events to a widget for scrolling its associated canvas."""
        canvas = None
        if isinstance(widget, tk.Canvas):
            canvas = widget
        else: # Find canvas associated with the scrollable frame
            for frame, info in self._scroll_widgets_by_tab.items():
                 if widget == frame or (hasattr(widget, 'master') and widget.master == frame):
                     canvas = info['canvas']
                     break

        if canvas and canvas.winfo_exists() and widget not in getattr(self, '_mousewheel_bound_widgets', set()):
            # Ensure set exists
            if not hasattr(self, '_mousewheel_bound_widgets'): self._mousewheel_bound_widgets = set()

            widget.bind("<MouseWheel>", lambda e, c=canvas: self._on_mousewheel(e, c), add='+')
            widget.bind("<Button-4>", lambda e, c=canvas: self._on_mousewheel(e, c), add='+')
            widget.bind("<Button-5>", lambda e, c=canvas: self._on_mousewheel(e, c), add='+')
            self._mousewheel_bound_widgets.add(widget)

    # IVO+GPT
    def _unbind_all_mousewheel(self):
        """Unbinds mousewheel events from all tracked widgets in the dialog."""
        if not hasattr(self, '_mousewheel_bound_widgets'): return
        widgets_to_unbind = list(self._mousewheel_bound_widgets)
        for widget in widgets_to_unbind:
             if widget and widget.winfo_exists():
                 try:
                     widget.unbind("<MouseWheel>")
                     widget.unbind("<Button-4>")
                     widget.unbind("<Button-5>")
                 except tk.TclError: pass
                 except Exception as e: logging.error(f"Error unbinding mousewheel from {widget}: {e}")
        self._mousewheel_bound_widgets.clear()

    # IVO+GPT
    def _on_mousewheel(self, event, canvas):
        """Handles mousewheel events for the specified canvas."""
        if not canvas or not canvas.winfo_exists(): return
        delta = 0
        scroll_multiplier = 2 # Sensitivity
        if event.num == 4: delta = -1 * scroll_multiplier
        elif event.num == 5: delta = 1 * scroll_multiplier
        elif hasattr(event, 'delta'):
            delta = -1 * scroll_multiplier if event.delta > 0 else (1 * scroll_multiplier if event.delta < 0 else 0)

        if delta != 0:
            try:
                scroll_info = canvas.yview()
                if (delta < 0 and scroll_info[0] > 0.0) or (delta > 0 and scroll_info[1] < 1.0):
                    canvas.yview_scroll(delta, "units")
            except tk.TclError as e: logging.warning(f"TclError during dialog canvas scroll: {e}")

    # IVO+GPT
    # --- Input Row Helper ---
    def _add_input_row(self, parent, row, label_text, var_dict, var_key, input_type="entry", values=None, unit=None, desc=None, required=False, **kwargs):
        """Helper to add a label, input widget, unit, and description row."""
        # Styles based on dialog theme
        label_style = 'DialogBold.TLabel' if required else 'Dialog.TLabel'
        entry_style = 'Dialog.TEntry'; combo_style = 'Dialog.TCombobox'
        # Use Card Checkbutton style as it matches the background
        check_style = 'Dialog.TCheckbutton'
        unit_style = 'Unit.TLabel'; desc_style = 'DialogDesc.TLabel'

        # Create tk variable
        initial_value = kwargs.get('initial')
        if input_type == "checkbutton":
            var = tk.BooleanVar(value=initial_value if isinstance(initial_value, bool) else False)
        else:
            var = tk.StringVar(value=str(initial_value) if initial_value is not None else "")
        var_dict[var_key] = var

        # Label
        label_full_text = label_text + (" *" if required else "")
        label = ttk.Label(parent, text=label_full_text, style=label_style)
        label.grid(row=row, column=0, sticky="nw", padx=(0, 10), pady=(5, 0))

        # Input Widget Frame (for widget + unit)
        widget_frame = tk.Frame(parent, bg=theme_colors[DLG_CARD])
        # Grid frame in second column, configure expansion
        widget_frame.grid(row=row, column=1, sticky="ew", padx=0, pady=(5, 0))
        widget_frame.grid_columnconfigure(0, weight=1) # Input widget expands

        widget = None
        if input_type == "entry":
            widget = ttk.Entry(widget_frame, textvariable=var, style=entry_style, width=kwargs.get('width', 25))
            widget.grid(row=0, column=0, sticky="ew")
        elif input_type == "combobox":
            widget = ttk.Combobox(widget_frame, textvariable=var, values=values or [], state='readonly', style=combo_style, width=kwargs.get('width', 23))
            # Set initial value robustly if provided and valid
            current_val = var.get()
            if values and (current_val not in values or not current_val):
                 if initial_value in values: widget.set(initial_value)
                 # elif values: widget.set(values[0]) # Optionally default to first item
            widget.grid(row=0, column=0, sticky="ew")
        elif input_type == "checkbutton":
            # Place checkbutton directly in parent grid (column 1), no frame/unit needed
            widget = ttk.Checkbutton(parent, variable=var, style=check_style, text="") # Use Dialog style
            widget.grid(row=row, column=1, sticky="w", padx=0, pady=(5,0))
            widget_frame.grid_forget() # Remove the now unused frame

        # Unit Label (if applicable and not checkbutton)
        if unit and input_type != "checkbutton":
            unit_label = ttk.Label(widget_frame, text=unit, style=unit_style)
            unit_label.grid(row=0, column=1, sticky="w", padx=(5, 0))
            widget_frame.grid_columnconfigure(1, weight=0) # Unit col takes no extra space
        elif input_type != "checkbutton":
             widget_frame.grid_columnconfigure(1, weight=0) # Reserve space even without unit

        # Description Label
        next_row = row + 1
        if desc:
            desc_label = ttk.Label(parent, text=desc, style=desc_style, wraplength=450) # Wider wrap
            desc_label.grid(row=row + 1, column=1, sticky="nw", padx=0, pady=(2, 10)) # Place below input col
            next_row = row + 2

        return widget, next_row


    # --- Tab Widget Creation Functions ---
    # (Residential, Travel, Food, GoodsWaste, Services, Digital)

    # IVO+GPT
    def create_residential_tab_widgets(self, scrollable_content_frame):
        vars_res = self.activity_vars["residential"]
        current_row = 0
        parent = scrollable_content_frame # Use the correct parent

        # --- Electricity Usage ---
        ttk.Label(parent, text="Grid Electricity", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        _, current_row = self._add_input_row(parent, current_row, "Electricity Used:", vars_res, "elec_kwh", unit="kWh", desc="Grid electricity consumption for the period.", required=True)
        _, current_row = self._add_input_row(parent, current_row, "Billing Period:", vars_res, "elec_period", "combobox", ["Monthly", "Bi-monthly", "Quarterly", "Annually"], desc="Duration covered by the electricity usage.", required=True, initial="Monthly")

        # --- Heating & Cooling ---
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Heating & Cooling", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        heat_fuel_types = ["None", "Natural Gas", "Heating Oil", "Propane", "Wood"]
        fuel_combo, current_row = self._add_input_row(parent, current_row, "Primary Fuel:", vars_res, "heat_fuel_type", "combobox", heat_fuel_types, desc="Main fuel used for heating/cooling (if any).", initial="None")

        # Conditional Fuel Frame (Spans 2 cols, placed below combo)
        self.res_heat_details_frame = tk.Frame(parent, bg=theme_colors[DLG_CARD])
        self.res_heat_details_frame.grid(row=current_row, column=0, columnspan=2, sticky='nsew', pady=(0, 5)) # Reduced padding
        self.res_heat_details_frame.grid_columnconfigure(1, weight=1)
        current_row += 1 # Next row in main parent

        self.res_heat_widgets = {} # Reset container
        if fuel_combo: fuel_combo.bind("<<ComboboxSelected>>", self._on_res_heat_fuel_change)
        self._on_res_heat_fuel_change() # Initial population

        # Heating Period (Now below the conditional frame)
        _, current_row = self._add_input_row(parent, current_row, "Fuel Period:", vars_res, "heat_fuel_period", "combobox", ["Monthly", "Quarterly", "Annually"], desc="Period for fuel amount (if entered).", initial="Monthly")


        # --- Water Heating ---
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Water Heating", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        water_heater_types = ["Electric", "Natural Gas", "Solar Thermal", "None"]
        water_type_combo, current_row = self._add_input_row(parent, current_row, "Heater Type:", vars_res, "water_heater_type", "combobox", water_heater_types, desc="Type of water heater used (if any).", initial="Electric")

        # Conditional Water Frame
        self.res_water_details_frame = tk.Frame(parent, bg=theme_colors[DLG_CARD])
        self.res_water_details_frame.grid(row=current_row, column=0, columnspan=2, sticky='nsew', pady=(0, 5))
        self.res_water_details_frame.grid_columnconfigure(1, weight=1)
        current_row += 1

        self.res_water_widgets = {} # Reset container
        if water_type_combo: water_type_combo.bind("<<ComboboxSelected>>", self._on_res_water_type_change)
        self._on_res_water_type_change()

        # Water Period (Now below conditional frame)
        _, current_row = self._add_input_row(parent, current_row, "Usage Period:", vars_res, "water_usage_period", "combobox", ["Monthly", "Quarterly", "Annually"], desc="Period for usage amount (if entered).", initial="Monthly")


        # --- On-Site Renewables ---
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="On-Site Renewables (Optional)", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        renew_types = ["None", "Solar Panels", "Wind Turbines"]
        _, current_row = self._add_input_row(parent, current_row, "Installation Type:", vars_res, "renew_type", "combobox", renew_types, initial="None")
        _, current_row = self._add_input_row(parent, current_row, "kWh Generated:", vars_res, "renew_kwh_gen", unit="kWh", desc="Energy generated by system (offsets emissions).")
        _, current_row = self._add_input_row(parent, current_row, "Generation Period:", vars_res, "renew_period", "combobox", ["Monthly", "Quarterly", "Annually"], desc="Period for kWh generated.", initial="Monthly")

    # IVO+GPT
    def _on_res_heat_fuel_change(self, event=None):
        """Shows/hides heating fuel amount/type widgets based on selection."""
        # Ensure instance vars are ready
        if not hasattr(self, 'activity_vars') or "residential" not in self.activity_vars or not hasattr(self, 'res_heat_details_frame'): return
        vars_res = self.activity_vars["residential"]
        selected_fuel = vars_res.get("heat_fuel_type").get() # Safely get var and its value
        parent_frame = self.res_heat_details_frame
        widgets = self.res_heat_widgets # Use instance attribute

        # Clear existing conditional widgets and associated variables
        for widget in list(parent_frame.winfo_children()): widget.destroy()
        widgets.clear()
        vars_res.pop("heat_fuel_amount", None)
        vars_res.pop("heat_wood_type", None)

        cond_row = 0
        unit_map = {"Natural Gas": "therms", "Heating Oil": "gallons", "Propane": "gallons", "Wood": "cords"}
        unit = unit_map.get(selected_fuel)

        if unit: # If a fuel requiring an amount is selected
            amount_label = "Wood Amount:" if selected_fuel == "Wood" else "Fuel Amount:"
            amount_widget, cond_row = self._add_input_row(parent_frame, cond_row, amount_label, vars_res, "heat_fuel_amount", unit=unit, desc=f"Amount of {selected_fuel.lower()} used.")
            widgets["amount"] = amount_widget

            if selected_fuel == "Wood":
                wood_widget, cond_row = self._add_input_row(parent_frame, cond_row, "Wood Type:", vars_res, "heat_wood_type", "combobox", ["Hardwood", "Softwood"], initial="Hardwood")
                widgets["wood_type"] = wood_widget

    # IVO+GPT
    def _on_res_water_type_change(self, event=None):
        """Shows/hides water heating usage widgets based on type."""
        if not hasattr(self, 'activity_vars') or "residential" not in self.activity_vars or not hasattr(self, 'res_water_details_frame'): return
        vars_res = self.activity_vars["residential"]
        selected_type = vars_res.get("water_heater_type").get()
        parent_frame = self.res_water_details_frame
        widgets = self.res_water_widgets # Use instance attribute

        for widget in list(parent_frame.winfo_children()): widget.destroy()
        widgets.clear()
        vars_res.pop("water_usage_amount", None)

        cond_row = 0
        unit_map = {"Electric": "kWh", "Natural Gas": "therms", "Solar Thermal": "kWh"}
        unit = unit_map.get(selected_type)

        if unit: # If type requires usage input
            usage_label = "Energy Used:"
            usage_widget, cond_row = self._add_input_row(parent_frame, cond_row, usage_label, vars_res, "water_usage_amount", unit=unit, desc=f"Energy used by {selected_type.lower()} heater.")
            widgets["usage"] = usage_widget

    # IVO+GPT
    def create_travel_tab_widgets(self, scrollable_content_frame):
        vars_travel = self.activity_vars["travel"]
        current_row = 0
        parent = scrollable_content_frame

        modes = ["Car", "Motorcycle", "Bus", "Train", "Subway", "Jeepney", "Air Travel", "Rideshare"]
        mode_combo, current_row = self._add_input_row(parent, current_row, "Mode:", vars_travel, "mode", "combobox", modes, required=True, initial="Car")
        self.travel_widgets["mode_combo"] = mode_combo

        _, current_row = self._add_input_row(parent, current_row, "Distance:", vars_travel, "distance", unit="km", desc="Total distance for the specified period.", required=True)
        period_options = ["Per Trip", "Daily Total", "Weekly Total", "Monthly Total"]
        _, current_row = self._add_input_row(parent, current_row, "Period/Frequency:", vars_travel, "period", "combobox", period_options, desc="Timeframe this distance represents.", required=True, initial="Per Trip")

        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        base_details_end_row = current_row

        # --- Conditional Sections Frame ---
        self.travel_widgets["conditional_frame"] = tk.Frame(parent, bg=theme_colors[DLG_CARD])
        self.travel_widgets["conditional_frame"].grid(row=base_details_end_row, column=0, columnspan=2, sticky='nsew')
        self.travel_widgets["conditional_frame"].grid_columnconfigure(1, weight=1)

        # --- Define Sub-Frames (don't grid yet) ---
        self.travel_widgets["car_details"] = tk.Frame(self.travel_widgets["conditional_frame"], bg=theme_colors[DLG_CARD])
        self.travel_widgets["car_details"].grid_columnconfigure(1, weight=1)
        self.travel_widgets["rideshare_details"] = tk.Frame(self.travel_widgets["conditional_frame"], bg=theme_colors[DLG_CARD])
        self.travel_widgets["rideshare_details"].grid_columnconfigure(1, weight=1)
        self.travel_widgets["air_details"] = tk.Frame(self.travel_widgets["conditional_frame"], bg=theme_colors[DLG_CARD])
        self.travel_widgets["air_details"].grid_columnconfigure(1, weight=1)

        if mode_combo: mode_combo.bind("<<ComboboxSelected>>", self._on_travel_mode_change)
        self._on_travel_mode_change() # Initial call

    # IVO+GPT
    def _on_travel_mode_change(self, event=None):
        """Shows/hides relevant travel detail frames."""
        # Ensure instance vars are ready
        if not hasattr(self, 'activity_vars') or "travel" not in self.activity_vars or not hasattr(self, 'travel_widgets'): return
        selected_mode = self.activity_vars["travel"].get("mode").get()
        logging.debug(f"Travel mode changed to: {selected_mode}")

        conditional_frame = self.travel_widgets.get("conditional_frame")
        car_details = self.travel_widgets.get("car_details")
        rideshare_details = self.travel_widgets.get("rideshare_details")
        air_details = self.travel_widgets.get("air_details")

        # Clear conditional frame content (ungrid sub-frames)
        if conditional_frame and conditional_frame.winfo_exists():
             for widget in list(conditional_frame.winfo_children()):
                  widget.grid_forget()

        # Clear previously relevant variables
        vars_travel = self.activity_vars["travel"]
        conditional_keys = ["car_fuel_type", "rideshare_fuel_type", "rideshare_passengers", "flight_type", "flight_cabin"]
        for key in conditional_keys: vars_travel.pop(key, None)

        # Show and populate the relevant sub-frame
        sub_frame_to_show = None
        if selected_mode == "Car":
            sub_frame_to_show = car_details
            if sub_frame_to_show:
                 for w in list(sub_frame_to_show.winfo_children()): w.destroy() # Clear previous
                 self._add_input_row(sub_frame_to_show, 0, "Fuel Type:", vars_travel, "car_fuel_type", "combobox", ["Gasoline", "Diesel", "Electric"], desc="Fuel type of your car.", initial="Gasoline", required=True)
        elif selected_mode == "Rideshare":
            sub_frame_to_show = rideshare_details
            if sub_frame_to_show:
                 for w in list(sub_frame_to_show.winfo_children()): w.destroy()
                 _, next_row = self._add_input_row(sub_frame_to_show, 0, "Fuel Type:", vars_travel, "rideshare_fuel_type", "combobox", ["Gasoline", "Diesel", "Electric"], desc="Fuel type of the vehicle.", initial="Gasoline", required=True)
                 self._add_input_row(sub_frame_to_show, next_row, "Passengers:", vars_travel, "rideshare_passengers", unit="# (incl. driver)", desc="Total people in vehicle.", required=True, initial="1")
        elif selected_mode == "Air Travel":
            sub_frame_to_show = air_details
            if sub_frame_to_show:
                 for w in list(sub_frame_to_show.winfo_children()): w.destroy()
                 flight_types = ["Short (<1500km)", "Medium (1500-6000km)", "Long (>6000km)"] # Simplified labels
                 _, next_row = self._add_input_row(sub_frame_to_show, 0, "Flight Type:", vars_travel, "flight_type", "combobox", flight_types, desc="One-way distance estimate.", initial="Short (<1500km)", required=True)
                 cabin_classes = ["Economy", "Business", "First"]
                 self._add_input_row(sub_frame_to_show, next_row, "Cabin Class:", vars_travel, "flight_cabin", "combobox", cabin_classes, desc="Your ticket class.", initial="Economy", required=True)

        # Grid the selected sub-frame inside the conditional frame
        if sub_frame_to_show and sub_frame_to_show.winfo_exists():
            sub_frame_to_show.grid(row=0, column=0, columnspan=2, sticky='nsew')
            logging.debug(f"Showing details for {selected_mode}.")
        else:
            logging.debug(f"No specific details for {selected_mode}.")

    # IVO+GPT
    def create_food_tab_widgets(self, scrollable_content_frame):
        vars_food = self.activity_vars["food"]
        current_row = 0
        parent = scrollable_content_frame

        ttk.Label(parent, text="Average Food Consumption (Estimate)", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1

        # Use instance attribute _food_inputs_labels for display
        for key, label in self._food_inputs_labels.items():
             _, current_row = self._add_input_row(parent, current_row, f"{label}:", vars_food, key, unit="kg", desc="Estimated amount consumed.")

        _, current_row = self._add_input_row(parent, current_row, "Consumption Period:", vars_food, "consumption_period", "combobox", ["Per Week", "Per Month"], desc="Timeframe for the amounts above.", required=True, initial="Per Week")

        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Food Habits (Optional Adjustments)", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1

        local_options = ["Low (<25% Local)", "Medium (25-75% Local)", "High (>75% Local)"]
        _, current_row = self._add_input_row(parent, current_row, "Locally Sourced:", vars_food, "local_sourcing", "combobox", local_options, desc="Estimate % of food sourced locally.", initial="Medium (25-75% Local)")
        _, current_row = self._add_input_row(parent, current_row, "Primarily Organic:", vars_food, "organic_preference", "checkbutton", desc="Check if majority of relevant foods bought are organic.")
        package_options = ["Minimal (Bulk, Loose)", "Average Mix", "Mostly Packaged"]
        _, current_row = self._add_input_row(parent, current_row, "Packaging Level:", vars_food, "packaging_level", "combobox", package_options, desc="Typical packaging of groceries.", initial="Average Mix")
        regions = ["Luzon", "Visayas", "Mindanao", "Unknown/Other"]
        _, current_row = self._add_input_row(parent, current_row, "Region:", vars_food, "region", "combobox", regions, desc="Primary region (for potential adjustments).", initial="Luzon")

        ttk.Label(parent, text="Note: Food footprint is complex. This provides an estimate based on consumption and habits.", style='DialogDesc.TLabel', wraplength=400).grid(row=current_row, column=0, columnspan=2, sticky='nw', pady=(15, 0)); current_row += 1

    # IVO+GPT
    def create_goods_waste_tab_widgets(self, scrollable_content_frame):
        vars_gw = self.activity_vars["shopping"] # Use 'shopping' key
        current_row = 0
        parent = scrollable_content_frame

        # Goods Consumption
        ttk.Label(parent, text="Spending on Goods (PHP)", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        ttk.Label(parent, text="(Factors applied per PHP using approx. USD rate)", style='DialogDesc.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 10)); current_row += 1

        # Use instance attribute _spending_cats_labels for display
        for key, label in self._spending_cats_labels.items():
            _, current_row = self._add_input_row(parent, current_row, label.replace(" Spending", ":"), vars_gw, key, unit="PHP", desc="Amount spent this period.")

        _, current_row = self._add_input_row(parent, current_row, "Spending Period:", vars_gw, "spending_period", "combobox", ["Monthly", "Quarterly", "Annually", "One-off Purchase"], desc="Timeframe spending covers.", initial="Monthly")
        regions_retail = ["Urban", "Rural", "Unknown"]
        _, current_row = self._add_input_row(parent, current_row, "Area Type:", vars_gw, "area_type_retail", "combobox", regions_retail, desc="Affects retail & waste factors.", initial="Urban", required=True)

        # Waste Management
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Waste Management", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1

        _, current_row = self._add_input_row(parent, current_row, "Waste Amount:", vars_gw, "waste_kg", unit="kg", desc="Estimated total household waste generated.")
        _, current_row = self._add_input_row(parent, current_row, "Waste Period:", vars_gw, "waste_period", "combobox", ["Per Week", "Per Month"], desc="Timeframe for waste amount.", initial="Per Week")
        disposal_methods = ["Landfill (Unknown Methane)", "Landfill (Low Methane)", "Landfill (Medium Methane)", "Landfill (High Methane)", "Incineration", "Mixed Recycling & Waste"]
        _, current_row = self._add_input_row(parent, current_row, "Primary Disposal:", vars_gw, "waste_disposal", "combobox", disposal_methods, desc="How most waste is handled.", initial="Landfill (Unknown Methane)")

    # IVO+GPT
    def create_services_tab_widgets(self, scrollable_content_frame):
        vars_serv = self.activity_vars["services"]
        current_row = 0
        parent = scrollable_content_frame

        ttk.Label(parent, text="Service Usage Estimates", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1

        _, current_row = self._add_input_row(parent, current_row, "Dry Cleaning:", vars_serv, "dry_cleaning_kg", unit="kg", desc="Approx. weight of garments dry cleaned.")
        _, current_row = self._add_input_row(parent, current_row, "Dry Cleaning Period:", vars_serv, "dry_cleaning_period", "combobox", ["Per Month", "Per Year"], initial="Per Month")
        _, current_row = self._add_input_row(parent, current_row, "Landscaping Area:", vars_serv, "landscaping_m2", unit="m²", desc="Area serviced by landscaping.")
        _, current_row = self._add_input_row(parent, current_row, "Landscaping Period:", vars_serv, "landscaping_period", "combobox", ["Per Month", "Per Year"], initial="Per Month")
        regions_serv = ["Urban", "Rural", "Unknown"]
        _, current_row = self._add_input_row(parent, current_row, "Area Type:", vars_serv, "area_type_services", "combobox", regions_serv, desc="Affects service emission factors.", initial="Urban", required=True)

    # IVO+GPT
    def create_digital_tab_widgets(self, scrollable_content_frame):
        vars_digital = self.activity_vars["digital"]
        current_row = 0
        parent = scrollable_content_frame

        # Device Usage
        ttk.Label(parent, text="Average Daily Device Usage", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        _, current_row = self._add_input_row(parent, current_row, "Laptop Use:", vars_digital, "laptop_hours", unit="hours/day", desc="Avg. daily active use.")
        _, current_row = self._add_input_row(parent, current_row, "Mobile Use:", vars_digital, "mobile_hours", unit="hours/day", desc="Avg. daily active use.")
        _, current_row = self._add_input_row(parent, current_row, "Tablet Use:", vars_digital, "tablet_hours", unit="hours/day", desc="Avg. daily active use.")

        # Streaming / Gaming
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Streaming & Gaming (Avg. Daily)", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        stream_quality = ["Low (SD)", "Medium (HD)", "High (4K)"]
        _, current_row = self._add_input_row(parent, current_row, "Streaming Quality:", vars_digital, "streaming_quality", "combobox", stream_quality, initial="Medium (HD)")
        _, current_row = self._add_input_row(parent, current_row, "Streaming Hours:", vars_digital, "streaming_hours", unit="hours/day")
        gaming_type = ["Low Demand", "High Demand"] # Simplified labels
        _, current_row = self._add_input_row(parent, current_row, "Gaming Type:", vars_digital, "gaming_type", "combobox", gaming_type, initial="Low Demand")
        _, current_row = self._add_input_row(parent, current_row, "Gaming Hours:", vars_digital, "gaming_hours", unit="hours/day")

        # Data Usage
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        ttk.Label(parent, text="Internet Data Usage", style='SectionHeader.TLabel').grid(row=current_row, column=0, columnspan=2, sticky='w', pady=(0, 5)); current_row += 1
        _, current_row = self._add_input_row(parent, current_row, "Total Data:", vars_digital, "data_usage_gb", unit="GB", desc="Estimate total monthly data (WiFi & Mobile).")
        _, current_row = self._add_input_row(parent, current_row, "Data Period:", vars_digital, "data_period", "combobox", ["Per Month", "Per Day"], desc="Timeframe for data usage.", initial="Per Month", required=True)

        # Regional Grid Factor
        ttk.Separator(parent, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=15); current_row += 1
        regions_grid = ["Luzon", "Visayas", "Mindanao", "Unknown/Default"]
        _, current_row = self._add_input_row(parent, current_row, "Region (Grid):", vars_digital, "region_grid", "combobox", regions_grid, desc="Select grid for electricity emission factor.", initial="Luzon", required=True)

    # IVO-ONLY
    # --- Input Conversion/Validation Helpers ---
    def _get_float_or_zero(self, value_str):
        """Safely convert string to float, returning 0.0 on failure or empty."""
        if value_str is None: return 0.0
        try:
            cleaned_str = str(value_str).strip()
            return float(cleaned_str) if cleaned_str else 0.0
        except (ValueError, TypeError):
            return 0.0

    # IVO-ONLY
    def _get_int_or_zero(self, value_str):
        """Safely convert string to int, returning 0 on failure or empty."""
        if value_str is None: return 0
        try:
             cleaned_str = str(value_str).strip()
             return int(float(cleaned_str)) if cleaned_str else 0
        except (ValueError, TypeError):
             return 0

    # IVO-ONLY
    # --- Carbon Calculation Logic ---
    def _calculate_carbon_footprint(self, category, details):
        """Calculates CO2e based on validated and cleaned input details."""
        # Use factors loaded into instance attribute
        factors = self.factors
        total_co2e = 0.0
        calculation_steps = [] # For logging/debugging

        # Pre-calculate average days/weeks per month
        DAYS_PER_MONTH = 30.4375 # More precise average
        WEEKS_PER_MONTH = DAYS_PER_MONTH / 7.0

        # IVO-ONLY
        # Helper to get monthly amount from period total
        def get_monthly_average(amount, period, months_map=None):
             # Treat "Per Trip" as a one-off contribution for this period's calculation
             # Treat "One-off Purchase" as averaged over a year
             if months_map is None:
                 months_map = {"Monthly": 1, "Bi-monthly": 2, "Quarterly": 3, "Annually": 12,
                               "Per Week": 1 / WEEKS_PER_MONTH, "Per Day": 1 / DAYS_PER_MONTH, "Per Trip": 1,
                               "One-off Purchase": 12}
             if period == "Monthly": return amount
             if period == "Per Week": return amount * WEEKS_PER_MONTH
             if period == "Daily Total": return amount * DAYS_PER_MONTH
             if period == "Annually": return amount / 12.0
             if period == "Quarterly": return amount / 3.0
             if period == "Bi-monthly": return amount / 2.0
             if period == "One-off Purchase": return amount / 12.0 # Average one-off over a year
             if period == "Per Trip": return amount # Treat trip as its own contribution

             # Fallback for unknown periods
             logging.warning(f"Unknown period '{period}' encountered in calculation. Using raw amount.")
             return amount


        try:
            # IVO-ONLY
            # --- Residential ---
            if category == "residential":
                # Elec
                elec_kwh = self._get_float_or_zero(details.get("elec_kwh"))
                elec_period = details.get("elec_period", "Monthly") # Should be validated
                monthly_elec = get_monthly_average(elec_kwh, elec_period)
                elec_fp = monthly_elec * factors.get("res_elec_usage_ph_nat_avg_kwh", 0)
                if monthly_elec > 0: calculation_steps.append(f"Elec: {monthly_elec:.1f}kWh/mo -> {elec_fp:.2f}")
                total_co2e += elec_fp

                # Heat
                heat_fuel = details.get("heat_fuel_type")
                heat_amount = self._get_float_or_zero(details.get("heat_fuel_amount"))
                heat_period = details.get("heat_fuel_period", "Monthly")
                monthly_heat = get_monthly_average(heat_amount, heat_period)
                heat_fp = 0.0
                if monthly_heat > 0 and heat_fuel != "None":
                    factor_key = None
                    if heat_fuel == "Natural Gas": factor_key = "res_heat_nat_gas_therm"
                    elif heat_fuel == "Heating Oil": factor_key = "res_heat_heating_oil_gallon"
                    elif heat_fuel == "Propane": factor_key = "res_heat_propane_gallon"
                    elif heat_fuel == "Wood":
                         wood_type = details.get("heat_wood_type", "Hardwood")
                         factor_key = "res_heat_wood_softwood_cord" if wood_type == "Softwood" else "res_heat_wood_hardwood_cord"
                    if factor_key: heat_fp = monthly_heat * factors.get(factor_key, 0)
                    calculation_steps.append(f"Heat ({heat_fuel}): {monthly_heat:.2f} units/mo -> {heat_fp:.2f}")
                total_co2e += heat_fp

                # Water Heat
                water_type = details.get("water_heater_type")
                water_usage = self._get_float_or_zero(details.get("water_usage_amount"))
                water_period = details.get("water_usage_period", "Monthly")
                monthly_water = get_monthly_average(water_usage, water_period)
                water_fp = 0.0
                if monthly_water > 0 and water_type != "None":
                    factor_key = None
                    if water_type == "Electric": factor_key = "res_water_elec_kwh"
                    elif water_type == "Natural Gas": factor_key = "res_water_gas_therm"
                    elif water_type == "Solar Thermal": factor_key = "res_water_solar_thermal_kwh"
                    if factor_key: water_fp = monthly_water * factors.get(factor_key, 0)
                    calculation_steps.append(f"Water ({water_type}): {monthly_water:.2f} units/mo -> {water_fp:.2f}")
                total_co2e += water_fp

                # Renewables (Savings)
                renew_type = details.get("renew_type")
                renew_gen = self._get_float_or_zero(details.get("renew_kwh_gen"))
                renew_period = details.get("renew_period", "Monthly")
                monthly_renew = get_monthly_average(renew_gen, renew_period)
                renew_fp = 0.0
                if monthly_renew > 0 and renew_type != "None":
                     factor_key = "res_renew_solar_panels_kwh" if renew_type == "Solar Panels" else "res_renew_wind_turbines_kwh"
                     renew_fp = monthly_renew * factors.get(factor_key, 0) # Factor is negative
                     calculation_steps.append(f"Renew ({renew_type}): {monthly_renew:.1f} kWh/mo -> {renew_fp:.2f}")
                total_co2e += renew_fp

                calculation_steps.append("Residential Result: Avg Monthly")

            # IVO-ONLY
            # --- Transportation ---
            elif category == "travel":
                mode = details.get("mode")
                distance_km = self._get_float_or_zero(details.get("distance"))
                period = details.get("period", "Per Trip") # Already validated

                # Get base factor (per km or pkm) and multipliers
                factor_val, is_pkm, multiplier, occupancy = 0.0, False, 1.0, 1
                if mode == "Car":
                    fuel = details.get("car_fuel_type"); f_key = {"Gasoline": "trans_pv_gasoline_km", "Diesel": "trans_pv_diesel_km", "Electric": "trans_pv_electric_km"}.get(fuel); factor_val = factors.get(f_key, 0); is_pkm = False
                elif mode == "Motorcycle": factor_val = factors.get("trans_pub_motorcycle_pkm", 0); is_pkm = True
                elif mode == "Bus": factor_val = factors.get("trans_pub_bus_pkm", 0); is_pkm = True
                elif mode == "Train": factor_val = factors.get("trans_pub_train_pkm", 0); is_pkm = True
                elif mode == "Subway": factor_val = factors.get("trans_pub_subway_pkm", 0); is_pkm = True
                elif mode == "Jeepney": factor_val = factors.get("trans_pub_jeepney_pkm", 0); is_pkm = True
                elif mode == "Air Travel":
                    flight_type = details.get("flight_type"); cabin = details.get("flight_cabin")
                    f_key = {"Short": "trans_air_short_pkm", "Medium": "trans_air_medium_pkm", "Long": "trans_air_long_pkm"}.get(flight_type.split()[0], "trans_air_medium_pkm") # Approx match
                    factor_val = factors.get(f_key, 0); is_pkm = True
                    m_key = {"Economy": "trans_air_cabin_economy", "Business": "trans_air_cabin_business", "First": "trans_air_cabin_first"}.get(cabin, "trans_air_cabin_economy")
                    multiplier = factors.get(m_key, 1.0)
                elif mode == "Rideshare":
                    fuel = details.get("rideshare_fuel_type"); passengers = self._get_int_or_zero(details.get("rideshare_passengers")); occupancy = max(1, passengers)
                    f_key = {"Gasoline": "trans_pv_gasoline_km", "Diesel": "trans_pv_diesel_km", "Electric": "trans_pv_electric_km"}.get(fuel); factor_val = factors.get(f_key, 0); is_pkm = False


                # Calculate footprint for the distance
                trip_fp = 0
                if is_pkm: trip_fp = distance_km * factor_val * multiplier
                else: trip_fp = (distance_km * factor_val * multiplier) / occupancy
                calculation_steps.append(f"Trip FP ({mode}): {trip_fp:.2f} (for {distance_km} km)")

                # Convert trip footprint to monthly average based on period
                total_co2e = get_monthly_average(trip_fp, period)
                calculation_steps.append(f"Travel Result: Avg Monthly (based on {period})")

            # IVO-ONLY
            # --- Food ---
            elif category == "food":
                consumption_period = details.get("consumption_period", "Per Week")
                monthly_prod_fp = 0.0
                total_monthly_kg = 0 # Track total kg for regional adjustments

                # Use instance attributes for mapping
                for input_key, factor_info in self._food_inputs_map.items():
                    amount_kg = self._get_float_or_zero(details.get(input_key))
                    if amount_kg <= 0: continue

                    monthly_kg = get_monthly_average(amount_kg, consumption_period)
                    total_monthly_kg += monthly_kg # Accumulate total for regional adjustment later

                    factor_val = 0
                    if isinstance(factor_info, tuple): # Average factors if tuple provided
                         f_vals = [factors.get(f_key, 0) for f_key in factor_info]
                         factor_val = sum(f_vals) / len(f_vals) if f_vals else 0
                    else: # Single factor key
                         factor_val = factors.get(factor_info, 0)

                    monthly_prod_fp += monthly_kg * factor_val
                calculation_steps.append(f"Production FP (Monthly Avg): {monthly_prod_fp:.2f}")

                # Apply Multipliers
                local = details.get("local_sourcing", "Medium"); organic = details.get("organic_preference", False); packaging = details.get("packaging_level", "Average")
                local_mult = {"Low": 1.05, "Medium": 1.0, "High": 0.90}.get(local.split()[0], 1.0)
                fert_conv = factors.get("food_farm_fertilizer_conventional_kgN", 1.0); fert_org = factors.get("food_farm_fertilizer_organic_kgN", 1.0); base_fert = (fert_conv + fert_org) / 2.0
                fert_mult = (fert_org / base_fert) if organic and base_fert > 0 else ((fert_conv / base_fert) if base_fert > 0 else 1.0)
                pkg_mult = {"Minimal": 0.95, "Average": 1.0, "Mostly": 1.10}.get(packaging.split()[0], 1.0)
                adjusted_fp = monthly_prod_fp * local_mult * fert_mult * pkg_mult
                calculation_steps.append(f"Adjusted FP (Qual: Loc*{local_mult:.2f}, Org*{fert_mult:.2f}, Pkg*{pkg_mult:.2f}): {adjusted_fp:.2f}")

                # Regional Additive Adjustment
                region_adj_fp = 0.0
                region = details.get("region", "Luzon")
                region_factor = 0.0
                if total_monthly_kg > 0:
                    region_map = {"Luzon": "food_region_luzon_kg_crop", "Visayas": "food_region_visayas_kg_crop", "Mindanao": "food_region_mindanao_kg_crop"}
                    region_factor = factors.get(region_map.get(region), 0)
                    if region_factor != 0:
                         region_adj_fp = total_monthly_kg * region_factor
                         calculation_steps.append(f"Region Adj ({region}): {total_monthly_kg:.1f}kg total/mo * {region_factor:.3f} = +{region_adj_fp:.2f}")

                total_co2e = adjusted_fp + region_adj_fp
                calculation_steps.append("Food Result: Avg Monthly")

            # IVO-ONLY
            # --- Goods & Waste (Was Shopping Before Waste Research was Unavailable) ---
            elif category == "shopping":
                # Spending
                spending_period = details.get("spending_period", "Monthly")
                area_type = details.get("area_type_retail", "Urban")
                period_spending_fp = 0.0
                usd_conv = 1.0 / PHP_TO_USD_RATE if PHP_TO_USD_RATE > 0 else 0

                # Use instance attributes for mapping
                for input_key, factor_key in self._spending_cats_map.items():
                     php_amount = self._get_float_or_zero(details.get(input_key))
                     if php_amount > 0 and usd_conv > 0:
                          usd_amount = php_amount * usd_conv
                          base_factor = factors.get(factor_key, 0)
                          period_spending_fp += usd_amount * base_factor

                region_mult = factors.get(f"goods_region_{area_type.lower()}_retail_mult", 1.0) if area_type != "Unknown" else 1.0
                period_spending_fp *= region_mult
                monthly_spending_fp = get_monthly_average(period_spending_fp, spending_period)
                calculation_steps.append(f"Spending FP (Monthly Avg, Region: {area_type}): {monthly_spending_fp:.2f}")

                # Waste
                waste_kg = self._get_float_or_zero(details.get("waste_kg"))
                monthly_waste_fp = 0.0
                if waste_kg > 0:
                    waste_period = details.get("waste_period", "Per Week")
                    disposal = details.get("waste_disposal", "Unknown")
                    monthly_waste_kg = get_monthly_average(waste_kg, waste_period)

                    waste_factor = 0.0
                    if "Recycling" in disposal: waste_factor = factors.get("waste_recycle_avg_mix_kg_kg", 0)
                    elif "Incineration" in disposal: waste_factor = factors.get("waste_incineration_kg_kg", 0)
                    else: # Landfill
                         lf_base_map = {"Low": "waste_landfill_low_ch4_kg_kg", "Medium": "waste_landfill_med_ch4_kg_kg", "High": "waste_landfill_high_ch4_kg_kg"}
                         lf_key = next((k for k in lf_base_map if k in disposal), "Medium") # Default Medium/Unknown
                         base_lf = factors.get(lf_base_map[lf_key], 0)
                         # Regional override
                         waste_factor = factors.get(f"waste_region_{area_type.lower()}_landfill_kg_kg", base_lf) if area_type != "Unknown" else base_lf
                    monthly_waste_fp = monthly_waste_kg * waste_factor
                    calculation_steps.append(f"Waste FP (Monthly Avg, Disposal: {disposal}): {monthly_waste_fp:.2f}")

                total_co2e = monthly_spending_fp + monthly_waste_fp
                calculation_steps.append("Goods & Waste Result: Avg Monthly")

            # IVO-ONLY
            # --- Services ---
            elif category == "services":
                area_type = details.get("area_type_services", "Urban")
                dc_kg = self._get_float_or_zero(details.get("dry_cleaning_kg"))
                ls_m2 = self._get_float_or_zero(details.get("landscaping_m2"))
                dc_period = details.get("dry_cleaning_period", "Per Month")
                ls_period = details.get("landscaping_period", "Per Month")

                # Dry Cleaning
                monthly_dc_fp = 0.0
                if dc_kg > 0:
                    monthly_dc_kg = get_monthly_average(dc_kg, dc_period)
                    f_key = f"serv_drycleaning_region_{area_type.lower()}_kg_garment"
                    dc_factor = factors.get(f_key, factors.get("serv_drycleaning_base_kg_garment", 0))
                    monthly_dc_fp = monthly_dc_kg * dc_factor
                    calculation_steps.append(f"Dry Cleaning ({area_type}): {monthly_dc_kg:.1f}kg/mo -> {monthly_dc_fp:.2f}")

                # Landscaping
                monthly_ls_fp = 0.0
                if ls_m2 > 0:
                    monthly_ls_m2 = get_monthly_average(ls_m2, ls_period)
                    f_key = f"serv_landscaping_region_{area_type.lower()}_m2"
                    ls_factor = factors.get(f_key, factors.get("serv_landscaping_base_m2", 0))
                    monthly_ls_fp = monthly_ls_m2 * ls_factor
                    calculation_steps.append(f"Landscaping ({area_type}): {monthly_ls_m2:.1f}m2/mo -> {monthly_ls_fp:.2f}")

                total_co2e = monthly_dc_fp + monthly_ls_fp
                calculation_steps.append("Services Result: Avg Monthly")

            # IVO-ONLY
            # --- Digital ---
            elif category == "digital":
                region = details.get("region_grid", "Luzon")
                grid_f_key = f"digital_grid_{region.lower()}_kwh"
                grid_factor = factors.get(grid_f_key, factors.get("digital_grid_default_kwh", 0))

                # Device energy (kWh/day)
                dev_kwh = (self._get_float_or_zero(details.get("laptop_hours")) * factors.get("digital_laptop_kwh_hour",0) +
                           self._get_float_or_zero(details.get("mobile_hours")) * factors.get("digital_mobile_kwh_hour",0) +
                           self._get_float_or_zero(details.get("tablet_hours")) * factors.get("digital_tablet_kwh_hour",0))

                # Stream/Game energy (kWh/day)
                sq = details.get("streaming_quality", "Medium"); sh = self._get_float_or_zero(details.get("streaming_hours"))
                gt = details.get("gaming_type", "Low"); gh = self._get_float_or_zero(details.get("gaming_hours"))
                stream_f = factors.get({"Low": "digital_stream_low_kwh_hour", "High": "digital_stream_high_kwh_hour"}.get(sq.split()[0], "digital_stream_medium_kwh_hour"), 0)
                game_f = factors.get({"Low": "digital_game_low_kwh_hour", "High": "digital_game_high_kwh_hour"}.get(gt.split()[0], "digital_game_low_kwh_hour"), 0)
                sg_kwh = (sh * stream_f) + (gh * game_f)

                # Data energy (kWh/day)
                data_gb = self._get_float_or_zero(details.get("data_usage_gb"))
                data_period = details.get("data_period", "Per Month")
                daily_data_gb = get_monthly_average(data_gb, data_period) / DAYS_PER_MONTH # Convert monthly avg GB to daily avg GB
                data_kwh_f = factors.get("digital_datacenter_kwh_gb", 0) + factors.get("digital_network_kwh_gb", 0)
                data_kwh = daily_data_gb * data_kwh_f

                # Total daily kWh and monthly CO2e
                total_kwh_day = dev_kwh + sg_kwh + data_kwh
                daily_co2e = total_kwh_day * grid_factor
                total_co2e = daily_co2e * DAYS_PER_MONTH
                calculation_steps.append(f"Digital: {total_kwh_day:.3f} kWh/day -> {daily_co2e:.3f} kgCO2e/day -> {total_co2e:.2f} kgCO2e/mo")
                calculation_steps.append("Digital Result: Avg Monthly")

            # Log calculation steps for debugging
            logging.debug(f"Calculation Steps for {category}:")
            for step in calculation_steps: logging.debug(f"  - {step}")

            # Final result: round and ensure non-negative
            final_co2e = round(max(0, total_co2e), 3)
            logging.info(f"Calculated footprint for {category}: {final_co2e:.3f} kg CO2e")
            return final_co2e

        # IVO+GPT
        except KeyError as e:
             logging.error(f"Missing factor key during calculation for {category}: {e}")
             messagebox.showerror("Calculation Error", f"Missing emission factor: '{e}'. Calculation failed.", parent=self)
             return None # Indicate failure
        except Exception as e:
            logging.exception(f"Unexpected error calculating footprint for {category}: {e}")
            messagebox.showerror("Calculation Error", f"Calculation error for {category}:\n{e}", parent=self)
            return None # Indicate failure


    # --- Submission Logic ---
    def submit_activity(self):
        """Gathers inputs, validates, calculates, saves, logs, and closes."""
        try:
            # IVO+GPT
            # 1. Get Active Category
            current_tab_id = self.notebook.select()
            if not current_tab_id: raise ValueError("No category tab selected.")
            container_frame = self.notebook.nametowidget(current_tab_id)

            # Find category key associated with the container frame's content frame
            activity_category = None
            for cat_key, scroll_frame in self.tabs.items():
                 info = self._scroll_widgets_by_tab.get(scroll_frame)
                 if info and info['canvas'].master == container_frame:
                      activity_category = cat_key
                      break
            if not activity_category:
                 # Fallback using tab text (less reliable)
                 try:
                     tab_text = self.notebook.tab(current_tab_id, "text")
                     for cat_key, cat_info in BASE_CATEGORIES.items():
                          if cat_info['name'] in tab_text:
                               activity_category = cat_key; break
                 except tk.TclError: pass
            if not activity_category: raise ValueError("Could not map tab to category.")

            logging.info(f"Submitting activity for category: {activity_category}")
            activity_vars = self.activity_vars[activity_category]
            activity_details = {key: var.get() for key, var in activity_vars.items()} # Raw inputs

            # IVO-ONLY
            # 2. Validate Inputs
            validation_errors = self._validate_inputs(activity_category, activity_details)
            if validation_errors:
                error_message = "Please correct the following:\n\n• " + "\n• ".join(validation_errors)
                messagebox.showerror("Invalid Input", error_message, parent=self)
                return

            # IVO-ONLY
            # 3. Clean Details (Convert numbers, handle Nones - for calculation)
            cleaned_details = self._clean_details_for_calculation(activity_details)

            # IVO-ONLY
            # 4. Calculate Footprint
            logging.debug(f"Calculating footprint for {activity_category} with cleaned details: {cleaned_details}")
            calculated_footprint = self._calculate_carbon_footprint(activity_category, cleaned_details)
            if calculated_footprint is None: return # Error already shown

            # IVO+GPT
            # 5. Create Record (Store RAW details for display, calculated FP)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_activity_record = {
                "timestamp": timestamp,
                "category": activity_category,
                "activity_details": activity_details, # Store raw user inputs
                "carbon_footprint": calculated_footprint
            }

            # IVO+GPT
            # 6. Update App State and Save
            if not isinstance(self.app_data.get("activities"), list): self.app_data["activities"] = []
            self.app_data["activities"].append(new_activity_record)
            log_activity(f"Added {activity_category.title()} activity ({calculated_footprint:.2f} kg CO₂e)")
            # Save ALL user data (includes new activity and log entry)
            if not save_user_data(self.app.current_user_id):
                # Attempt to rollback? Difficult. Best to inform user save failed.
                self.app_data["activities"].pop() # Remove activity from memory if save failed
                # How to remove last log entry? Complex. Leave as is for now.
                messagebox.showerror("Save Error", "Failed to save the new activity. Please try again.", parent=self)
                return

            # IVO+GPT
            # 7. Refresh Main UI & Close Dialog
            self.app.refresh_current_page()
            self.destroy()

        # IVO+GPT
        except ValueError as e: # Catch validation or internal logic errors
            messagebox.showerror("Input Error", str(e), parent=self)
            logging.error(f"Submission validation error: {e}", exc_info=True)
        except Exception as e:
            logging.exception("Unexpected error during activity submission")
            messagebox.showerror("Submission Error", f"An unexpected error occurred:\n{e}", parent=self)

    # IVO-ONLY
    def _validate_inputs(self, category, details):
        """Performs validation based on category and returns list of error messages."""
        errors = []
        def check_num(key, name, allow_zero=True, is_int=False):
            val_str = str(details.get(key, "")).strip()
            is_valid = False
            if val_str:
                try:
                    val = float(val_str)
                    if (allow_zero and val >= 0) or (not allow_zero and val > 0):
                         if is_int and val != int(val): # Check if float represents int if needed
                              is_valid = False
                         else:
                              is_valid = True
                except ValueError: pass
            elif allow_zero: is_valid = True # Empty allowed if zero allowed
            if not is_valid:
                 req = f"{'positive' if not allow_zero else 'non-negative'}{' whole' if is_int else ''}"
                 errors.append(f"'{name}' must be a {req} number or empty.")

        # IVO-ONLY
        def check_req(key, name):
            val = details.get(key)
            if val is None or str(val).strip() == "" or str(val) == "None":
                 errors.append(f"'{name}' is required.")

        # IVO+GPT (REPEATING RULES)
        # --- Category-Specific Rules ---
        if category == "residential":
            check_req("elec_kwh", "Electricity Used"); check_num("elec_kwh", "Electricity Used", allow_zero=False)
            check_req("elec_period", "Electricity Billing Period")
            heat_fuel = details.get("heat_fuel_type"); heat_amt = self._get_float_or_zero(details.get("heat_fuel_amount"))
            if heat_fuel != "None": check_num("heat_fuel_amount", f"{heat_fuel} Amount"); # Allow 0
            water_type = details.get("water_heater_type"); water_amt = self._get_float_or_zero(details.get("water_usage_amount"))
            if water_type != "None": check_num("water_usage_amount", f"{water_type} Usage"); # Allow 0
            renew_type = details.get("renew_type"); renew_amt = self._get_float_or_zero(details.get("renew_kwh_gen"))
            if renew_type != "None": check_num("renew_kwh_gen", "kWh Generated"); # Allow 0

        elif category == "travel":
            check_req("mode", "Mode"); check_req("distance", "Distance"); check_num("distance", "Distance", allow_zero=False)
            check_req("period", "Period/Frequency")
            mode = details.get("mode")
            if mode == "Car": check_req("car_fuel_type", "Car Fuel")
            if mode == "Rideshare": check_req("rideshare_fuel_type", "Rideshare Fuel"); check_req("rideshare_passengers", "Passengers"); check_num("rideshare_passengers", "Passengers", allow_zero=False, is_int=True)
            if mode == "Air Travel": check_req("flight_type", "Flight Type"); check_req("flight_cabin", "Cabin")

        elif category == "food":
            check_req("consumption_period", "Consumption Period"); check_req("region", "Region")
            has_amount = any(self._get_float_or_zero(details.get(k)) > 0 for k in self._food_inputs_map)
            if not has_amount: errors.append("Enter amount for at least one food category.")
            else: # Only check positivity if amounts entered
                 for k, label in self._food_inputs_labels.items(): check_num(k, label)

        elif category == "shopping": # Goods & Waste
            has_spend = any(self._get_float_or_zero(details.get(k)) > 0 for k in self._spending_cats_map)
            has_waste = self._get_float_or_zero(details.get("waste_kg")) > 0
            if not has_spend and not has_waste: errors.append("Enter Spending or Waste details.")
            if has_spend or has_waste: check_req("area_type_retail", "Area Type")
            if has_spend: check_req("spending_period", "Spending Period"); # Check amounts are positive
            for k, label in self._spending_cats_labels.items(): check_num(k, label)
            if has_waste: check_req("waste_kg", "Waste Amount"); check_num("waste_kg", "Waste Amount", allow_zero=False); check_req("waste_period", "Waste Period"); check_req("waste_disposal", "Disposal Method")

        elif category == "services":
             has_dc = self._get_float_or_zero(details.get("dry_cleaning_kg")) > 0
             has_ls = self._get_float_or_zero(details.get("landscaping_m2")) > 0
             if not has_dc and not has_ls: errors.append("Enter details for Dry Cleaning or Landscaping.")
             if has_dc or has_ls: check_req("area_type_services", "Area Type")
             if has_dc: check_req("dry_cleaning_kg", "Dry Cleaning Amount"); check_num("dry_cleaning_kg", "Dry Cleaning Amount", allow_zero=False); check_req("dry_cleaning_period", "DC Period")
             if has_ls: check_req("landscaping_m2", "Landscaping Area"); check_num("landscaping_m2", "Landscaping Area", allow_zero=False); check_req("landscaping_period", "LS Period")

        elif category == "digital":
             usage_keys = ["laptop_hours", "mobile_hours", "tablet_hours", "streaming_hours", "gaming_hours", "data_usage_gb"]
             has_usage = any(self._get_float_or_zero(details.get(k)) > 0 for k in usage_keys)
             if not has_usage: errors.append("Enter usage for at least one digital activity.")
             else: # Check numbers are valid if entered
                 for k in usage_keys: check_num(k, k.replace("_", " ").title())
             if self._get_float_or_zero(details.get("data_usage_gb")) > 0: check_req("data_period", "Data Period")
             check_req("region_grid", "Region (Grid)")

        return errors

    # IVO-ONLY
    def _clean_details_for_calculation(self, raw_details):
         """Converts numeric strings to numbers, keeps others as is."""
         cleaned = {}
         for key, value in raw_details.items():
              if isinstance(value, (bool, int, float)):
                   cleaned[key] = value
              elif value is None:
                   cleaned[key] = None
              else: # Should be string or string-like from tkVar.get()
                   val_str = str(value).strip()
                   if val_str == "" or val_str == "None":
                        cleaned[key] = None
                   else:
                        try: # Attempt numeric conversion
                             num_val = float(val_str)
                             # Store as int if it's effectively whole
                             cleaned[key] = int(num_val) if num_val.is_integer() else num_val
                        except ValueError:
                             cleaned[key] = val_str # Keep as string if not numeric
         return cleaned

    # EXPENSEWISE
    def destroy(self):
        """Overrides destroy for dialog cleanup."""
        logging.debug("Destroying AddCarbonFootprintActivityDialog, cleaning up...")
        try:
            # Unbind notebook event
            if hasattr(self, 'notebook') and self.notebook and self.notebook.winfo_exists():
                try: self.notebook.unbind("<<NotebookTabChanged>>")
                except tk.TclError: pass

            # Unbind Escape key
            try: self.unbind("<Escape>")
            except tk.TclError: pass

            # Unbind mousewheel events
            self._unbind_all_mousewheel()

        except Exception as e:
             logging.error(f"Error during AddCarbonFootprintActivityDialog cleanup: {e}", exc_info=True)
        finally:
            super().destroy() # Call original destroy

# --- Simple Entry Dialog Class (Used for Profile Creation) ---
# This class seems fine as is, already relatively simple. Ensure theme handling works.
class SimpleEntryDialog(tk.Toplevel):
    """Simple modal dialog for single text input (e.g., profile name)."""

    # EXPENSEWISE
    def __init__(self, parent, title, fields_config):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.title(title)
        self.fields_config = fields_config # Expects {"field_name": {"label": "...", "required": True/False}}
        self.result = None
        self.entries = {}
        self.vars = {}

        # Determine theme based on parent type or app instance
        # Default to dark if parent unknown or not App/AccountsPage
        dialog_theme = THEME_ECO_DARK.copy()
        app_instance = None
        curr = parent
        while curr:
            if isinstance(curr, ECOHUBApp): app_instance = curr; break
            if hasattr(curr, 'app') and isinstance(curr.app, ECOHUBApp): app_instance = curr.app; break
            if hasattr(curr, 'master'): curr = curr.master
            else: break
        if app_instance:
             current_app_theme = getattr(app_instance, 'current_theme', 'eco_dark')
             dialog_theme = THEME_ECO_LIGHT.copy() if current_app_theme == 'eco_light' else THEME_ECO_DARK.copy()
        elif isinstance(parent, AccountsPage): # Accounts page is always dark
             dialog_theme = THEME_ECO_DARK.copy()

        # Configure dialog appearance using determined theme
        dlg_bg = dialog_theme.get(DLG_BG, "#2d3347"); dlg_fg = dialog_theme.get(DLG_FG, "#c5cae9")
        card_bg = dialog_theme.get(DLG_CARD, "#353c50"); accent = dialog_theme.get(ACCENT, "#7986cb")
        btn_fg = dialog_theme.get(BTN_FG, "#ffffff"); disabled_bg = dialog_theme.get(DISABLED, "#607d8b")
        err_red = dialog_theme.get(RED, "#f44336"); accent_darker = dialog_theme.get(ACCENT_DARKER, accent)

        self.configure(bg=dlg_bg, padx=20, pady=20)
        main_frame = tk.Frame(self, bg=dlg_bg)
        main_frame.pack(expand=True, fill="both")
        main_frame.grid_columnconfigure(1, weight=1)

        # Dialog Styles (scoped to this instance)
        dialog_style = ttk.Style(self)
        dialog_style.theme_use('clam')
        dialog_style.configure('SimpleDialog.TLabel', background=dlg_bg, foreground=dlg_fg, font=FONT_NORMAL)
        dialog_style.configure('SimpleDialogBold.TLabel', background=dlg_bg, foreground=dlg_fg, font=FONT_BOLD)
        dialog_style.configure('SimpleDialog.TEntry', fieldbackground=card_bg, foreground=dlg_fg, insertcolor=dlg_fg, borderwidth=1, relief=tk.SOLID, bordercolor=disabled_bg, padding=5)
        dialog_style.map('SimpleDialog.TEntry', bordercolor=[('focus', accent)])
        dialog_style.configure('SimpleDialog.TButton', background=accent, foreground=btn_fg, font=FONT_BOLD, padding=(10, 5))
        dialog_style.map('SimpleDialog.TButton', background=[('active', accent_darker)])
        dialog_style.configure('Cancel.SimpleDialog.TButton', background=disabled_bg, foreground=dlg_fg, font=FONT_BOLD, padding=(10, 5))
        dialog_style.map('Cancel.SimpleDialog.TButton', background=[('active', err_red)], foreground=[('active', btn_fg)])

        # Create Input Fields
        row_num = 0
        for name, config in fields_config.items():
            label_text = config.get("label", name.replace("_", " ").title() + ":")
            required = config.get("required", False)
            lbl_style = 'SimpleDialogBold.TLabel' if required else 'SimpleDialog.TLabel'
            lbl = ttk.Label(main_frame, text=label_text + (" *" if required else ""), style=lbl_style)
            lbl.grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=5)

            var = tk.StringVar()
            entry = ttk.Entry(main_frame, textvariable=var, font=FONT_NORMAL, style='SimpleDialog.TEntry')
            entry.grid(row=row_num, column=1, sticky="ew", pady=5)
            self.entries[name] = entry
            self.vars[name] = var
            row_num += 1

        # Buttons
        button_frame = tk.Frame(main_frame, bg=dlg_bg)
        button_frame.grid(row=row_num, column=0, columnspan=2, sticky="e", pady=(15, 0))
        ok_btn = ttk.Button(button_frame, text="OK", command=self.on_ok, style="SimpleDialog.TButton")
        ok_btn.pack(side=tk.RIGHT, padx=(5, 0))
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.on_cancel, style="Cancel.SimpleDialog.TButton")
        cancel_btn.pack(side=tk.RIGHT)

        # Bindings and final setup
        self.bind("<Return>", self.on_ok)
        self.bind("<Escape>", self.on_cancel)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.after(100, self._set_initial_focus) # Delay focus setting slightly
        self._center_dialog()
        self.grab_set()
        self.wait_window(self)

    # EXPENSEWISE
    def _set_initial_focus(self):
        """Sets focus to the first Entry widget."""
        first_entry = next(iter(self.entries.values()), None) # Get first entry
        if first_entry:
            first_entry.focus_set()
            first_entry.select_range(0, tk.END)

    # EXPENSEWISE
    def _center_dialog(self):
        """Centers the dialog relative to its parent."""
        try:
            self.update_idletasks()
            if not self.parent or not self.parent.winfo_exists(): return
            parent_x, parent_y = self.parent.winfo_rootx(), self.parent.winfo_rooty()
            parent_w, parent_h = self.parent.winfo_width(), self.parent.winfo_height()
            dialog_w, dialog_h = self.winfo_width(), self.winfo_height()
            if parent_w <=0 or parent_h <=0 or dialog_w <=0 or dialog_h <=0: return

            x = parent_x + (parent_w // 2) - (dialog_w // 2)
            y = parent_y + (parent_h // 3) - (dialog_h // 2)
            screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight(); margin=10
            x = max(margin, min(x, screen_w - dialog_w - margin)); y = max(margin, min(y, screen_h - dialog_h - margin))
            self.geometry(f"+{x}+{y}")
        except Exception as e: logging.error(f"Error centering SimpleEntryDialog: {e}")

    # EXPENSEWISE
    def validate_required(self):
        """Checks required fields."""
        errors = []
        for name, config in self.fields_config.items():
             if config.get("required"):
                  if not self.vars.get(name).get().strip():
                       label = config.get("label", name).replace(":","").strip()
                       errors.append(f"'{label}' is required.")
        return errors

    # EXPENSEWISE
    def on_ok(self, event=None):
        """Handles OK click."""
        errors = self.validate_required()
        if errors:
             messagebox.showerror("Missing Information", "\n".join(errors), parent=self)
             return
        self.result = {name: var.get() for name, var in self.vars.items()}
        logging.debug(f"SimpleDialog OK. Result: {self.result}")
        self.destroy()

    # EXPENSEWISE
    def on_cancel(self, event=None):
        """Handles Cancel or close."""
        self.result = None
        logging.debug("SimpleDialog Cancelled.")
        self.destroy()

# IVO-ONLY (EXPENSEWISE ARCHITECTURE)
# --- Main Execution Logic ---
def launch_main_app(user_id):
    """Initializes and runs the main ECOHUBApp."""
    logging.info(f"Launching Main App for user: {user_id}...")
    app = None
    should_return_to_accounts = True # Default behavior

    try:
        app = ECOHUBApp(user_id)
        app.mainloop() # Blocks until window is closed

        # After mainloop ends, check if full exit was requested
        full_exit_req = getattr(app, '_full_exit_requested', False) if app else False
        if full_exit_req:
            logging.info("Full exit requested. Terminating.")
            should_return_to_accounts = False
        else:
             logging.info("Main app closed. Returning to accounts.")
             should_return_to_accounts = True

    except Exception as e:
        logging.exception(f"CRITICAL ERROR during main application execution for {user_id}")
        messagebox.showerror("Application Error", f"A critical error occurred:\n{e}\n\nReturning to profile selection.", parent=None)
        should_return_to_accounts = True # Attempt recovery

    return should_return_to_accounts

# IVO-ONLY (EXPENSEWISE ARCHITECTURE)
if __name__ == "__main__":
    # Perform initial setup outside the loop
    ensure_data_dir()
    # load_emission_factors() # Load factors once at startup? No, load per user session in load_user_data.
    logging.info("--- ECOHUB Application Starting ---")

    continue_running = True
    while continue_running:
        logging.info("Showing Accounts Page...")
        selected_user_id = None
        accounts_page = None

        try:
            accounts_page = AccountsPage()
            accounts_page.mainloop() # Blocks until AccountsPage closes

            # Check the result after the window is destroyed
            selected_user_id = getattr(accounts_page, 'selected_user_id', None)

            if selected_user_id is None:
                # User closed AccountsPage or clicked Exit
                logging.info("No user selected or Exit clicked. Shutting down.")
                continue_running = False
            else:
                # User selected, launch the main app
                # launch_main_app returns True if we should loop back to accounts
                continue_running = launch_main_app(selected_user_id)

        except Exception as e:
            logging.exception("FATAL ERROR during application startup or main loop")
            messagebox.showerror("Fatal Error", f"A critical error occurred:\n{e}\n\nThe application will close.", parent=None)
            continue_running = False # Stop on critical errors

    logging.info("--- ECOHUB Application Finished ---")