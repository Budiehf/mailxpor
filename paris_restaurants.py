import time
from typing import Dict, Iterable, List

import googlemaps
import pandas as pd

# --- Setup ---
API_KEY = "REPLACE_WITH_YOUR_API_KEY"
gmaps = googlemaps.Client(key=API_KEY)

barrios = [
    "Saint-Germain-l’Auxerrois",
    "Les Halles",
    "Palais-Royal",
    "Place-Vendôme",
    "Gaillon",
    "Vivienne",
    "Mail",
    "Bonne-Nouvelle",
    "Arts-et-Métiers",
    "Enfants-Rouges",
    "Archives",
    "Sainte-Avoye",
    "Saint-Merri",
    "Saint-Gervais",
    "Arsenal",
    "Notre-Dame",
    "Saint-Victor",
    "Jardin-des-Plantes",
    "Val-de-Grâce",
    "Sorbonne",
    "Monnaie",
    "Odéon",
    "Notre-Dame-des-Champs",
    "Saint-Germain-des-Prés",
    "Saint-Thomas-d’Aquin",
    "Invalides",
    "École-Militaire",
    "Gros-Caillou",
    "Champs-Élysées",
    "Faubourg-du-Roule",
    "Madeleine",
    "Europe",
    "Saint-Georges",
    "Chaussée-d’Antin",
    "Faubourg-Montmartre",
    "Rochechouart",
    "Saint-Vincent-de-Paul",
    "Porte-Saint-Denis",
    "Porte-Saint-Martin",
    "Hôpital-Saint-Louis",
    "Folie-Méricourt",
    "Saint-Ambroise",
    "La Roquette",
    "Sainte-Marguerite",
    "Bel-Air",
    "Picpus",
    "Bercy",
    "Quinze-Vingts",
    "Salpêtrière",
    "Gare",
    "Maison-Blanche",
    "Croulebarbe",
    "Montparnasse",
    "Parc-de-Montsouris",
    "Petit-Montrouge",
    "Plaisance",
    "Saint-Lambert",
    "Necker",
    "Grenelle",
    "Javel",
    "Auteuil",
    "Muette",
    "Porte-Dauphine",
    "Chaillot",
    "Ternes",
    "Plaine-de-Monceaux",
    "Batignolles",
    "Épinettes",
    "Grandes-Carrières",
    "Clignancourt",
    "Goutte-d’Or",
    "La Chapelle",
    "Villette",
    "Pont-de-Flandre",
    "Amérique",
    "Combat",
    "Belleville",
    "Saint-Fargeau",
    "Père-Lachaise",
    "Charonne",
]

restaurants: List[Dict[str, object]] = []


def push_items(items: Iterable[Dict[str, object]], tag: str) -> None:
    for it in items:
        restaurants.append(
            {
                "tag": tag,
                "name": it.get("name"),
                "place_id": it.get("place_id"),
                "price_level": it.get("price_level"),
                "rating": it.get("rating"),
                "user_ratings_total": it.get("user_ratings_total") or 0,
                "address": it.get("vicinity") or it.get("formatted_address"),
            }
        )


def drain_pagination(first_results: Dict[str, object], fetch_next: Dict[str, object]) -> None:
    results = first_results
    push_items(results.get("results", []), fetch_next["tag"])
    while results.get("next_page_token"):
        time.sleep(3)  # give token time to activate
        results = fetch_next["fn"](page_token=results["next_page_token"])
        push_items(results.get("results", []), fetch_next["tag"])


# --- Pass 1: by “barrio” around Paris (France) ---
for barrio in barrios:
    geo = gmaps.geocode(f"{barrio}, Paris, France")
    if not geo:
        continue
    loc = geo[0]["geometry"]["location"]
    lat, lng = loc["lat"], loc["lng"]
    first = gmaps.places_nearby(location=(lat, lng), radius=2000, type="restaurant")

    def fetch_next_page(**kw):
        return gmaps.places_nearby(location=(lat, lng), radius=2000, type="restaurant", **kw)

    drain_pagination(first, {"fn": fetch_next_page, "tag": barrio})


# --- Pass 2: by cuisine strings (Text Search) ---
cuisines = [
    "traditional",
    "mexican",
    "latin",
    "burger",
    "pizza",
    "high cuisine",
    "american",
    "italian",
    "turkish",
    "mediterranean",
    "chinese",
    "asian",
    "indian",
    "japanese",
    "french",
    "spanish",
    "vegetarian",
]

city = gmaps.geocode("Paris, France")[0]["geometry"]["location"]
v_lat, v_lng = city["lat"], city["lng"]

for cuisine in cuisines:
    query = f"{cuisine} restaurants in Paris, France"
    first = gmaps.places(query=query, location=(v_lat, v_lng), type="restaurant")

    def fetch_next_page(**kw):
        return gmaps.places(query=query, location=(v_lat, v_lng), type="restaurant", **kw)

    drain_pagination(first, {"fn": fetch_next_page, "tag": cuisine})


# --- Data shaping ---
rest_df = pd.DataFrame(restaurants).drop_duplicates(subset=["place_id"])
rest_df["user_ratings_total"] = rest_df["user_ratings_total"].fillna(0).astype(int)
rest_df["rating"] = pd.to_numeric(rest_df["rating"], errors="coerce")

filtered_df = rest_df.dropna(subset=["rating"])
filtered_df = filtered_df[filtered_df["rating"] >= 4.7]
filtered_df = filtered_df.sort_values(["rating", "user_ratings_total"], ascending=[False, False])

output_columns = [
    "name",
    "rating",
    "user_ratings_total",
    "price_level",
    "address",
    "tag",
]
filtered_df = filtered_df[output_columns]
filtered_df = filtered_df.rename(
    columns={
        "name": "Name",
        "rating": "Rating",
        "user_ratings_total": "User Ratings Total",
        "price_level": "Price Level",
        "address": "Address",
        "tag": "Source Tag",
    }
)

filtered_df.to_excel("paris_restaurants_4_7_plus.xlsx", index=False)
print(
    f"Saved {len(filtered_df)} restaurants with rating >= 4.7 to paris_restaurants_4_7_plus.xlsx"
)
