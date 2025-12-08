"""
Brand Database for Value Engineering
Contains furniture brands across different budget tiers with web scraping capabilities
"""

import requests
from bs4 import BeautifulSoup
import json
from typing import Dict, List, Optional
import logging

logger = logging.getLogger(__name__)


class BrandDatabase:
    """Database of furniture brands with web scraping capabilities"""
    
    def __init__(self):
        self.brands = self._initialize_brands()
    
    def _initialize_brands(self) -> Dict:
        """Initialize the furniture brand database"""
        return {
            'budgetary': {
                'seating': [
                    {
                        'name': 'Goldsit',
                        'website': 'https://www.goldsit.com',
                        'country': 'Turkey',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Goldsit Defy', 'price_range': '150-250', 'features': ['Mesh back', 'Lumbar support', 'Adjustable armrests']},
                                {'model': 'Goldsit Astro', 'price_range': '180-280', 'features': ['Leather option', 'Tilt mechanism', 'Gas lift']},
                            ],
                            'task_chairs': [
                                {'model': 'Goldsit Smart', 'price_range': '100-180', 'features': ['Mesh back', 'Basic adjustment', 'Swivel base']},
                                {'model': 'Goldsit Focus', 'price_range': '120-200', 'features': ['Ergonomic design', 'Height adjustment', 'Breathable fabric']},
                            ],
                            'visitor_chairs': [
                                {'model': 'Goldsit Slim', 'price_range': '80-120', 'features': ['Stackable', 'Chrome frame', 'Fixed armrests']},
                                {'model': 'Goldsit Neo', 'price_range': '90-140', 'features': ['Upholstered', 'Cantilever base', 'Comfortable padding']},
                            ],
                        }
                    },
                    {
                        'name': 'Kinwai',
                        'website': 'https://www.kinwai.com',
                        'country': 'China',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Kinwai Executive Pro', 'price_range': '140-230', 'features': ['PU leather', 'High back', 'Synchronized mechanism']},
                                {'model': 'Kinwai Elite', 'price_range': '160-250', 'features': ['Padded armrests', 'Recline function', 'Lumbar support']},
                            ],
                            'task_chairs': [
                                {'model': 'Kinwai Work', 'price_range': '90-160', 'features': ['Mesh back', 'Basic ergonomics', 'Adjustable height']},
                                {'model': 'Kinwai Office', 'price_range': '110-180', 'features': ['Multi-adjustment', 'Breathable mesh', 'Tilt lock']},
                            ],
                            'conference_chairs': [
                                {'model': 'Kinwai Meet', 'price_range': '120-180', 'features': ['Fixed armrests', 'Chrome frame', 'Upholstered seat']},
                            ],
                        }
                    },
                    {
                        'name': 'Malazan',
                        'website': 'https://www.malazan.com.my',
                        'country': 'Malaysia',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Malazan Director', 'price_range': '180-270', 'features': ['Genuine leather option', 'Ergonomic', 'Premium finish']},
                            ],
                            'task_chairs': [
                                {'model': 'Malazan Task Pro', 'price_range': '100-170', 'features': ['Mesh back', 'Adjustable armrests', 'Tilt mechanism']},
                                {'model': 'Malazan Lite', 'price_range': '80-140', 'features': ['Basic ergonomics', 'Fixed armrests', 'Gas lift']},
                            ],
                        }
                    },
                    {
                        'name': 'Sunon',
                        'website': 'https://www.sunon.com.tw',
                        'country': 'Taiwan',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Sunon Executive', 'price_range': '170-260', 'features': ['High density foam', 'Adjustable lumbar', 'Tilt mechanism']},
                            ],
                            'task_chairs': [
                                {'model': 'Sunon Ergo', 'price_range': '110-190', 'features': ['Mesh back', 'Adjustable headrest', 'Multi-function']},
                            ],
                            'visitor_chairs': [
                                {'model': 'Sunon Guest', 'price_range': '85-130', 'features': ['Stackable', 'Chrome legs', 'Upholstered']},
                            ],
                        }
                    },
                    {
                        'name': 'M&W',
                        'website': 'https://www.mw-furniture.com',
                        'country': 'China',
                        'models': {
                            'task_chairs': [
                                {'model': 'M&W Office Plus', 'price_range': '95-165', 'features': ['Mesh back', 'Ergonomic design', 'Basic adjustment']},
                            ],
                            'visitor_chairs': [
                                {'model': 'M&W Guest', 'price_range': '75-120', 'features': ['Stackable', 'Fixed armrests', 'Chrome frame']},
                            ],
                        }
                    },
                    {
                        'name': 'Merry Fair',
                        'website': 'https://www.merryfair.com',
                        'country': 'Taiwan',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Merry Fair Vito', 'price_range': '200-300', 'features': ['Mesh back', 'Premium ergonomics', 'Multi-adjustment']},
                            ],
                            'task_chairs': [
                                {'model': 'Merry Fair Deco', 'price_range': '130-210', 'features': ['Mesh', 'Ergonomic support', 'Adjustable armrests']},
                                {'model': 'Merry Fair Wau', 'price_range': '110-180', 'features': ['Mesh back', 'Basic ergonomics', 'Tilt mechanism']},
                            ],
                        }
                    },
                ],
                'desking': [
                    {
                        'name': 'Goldsit',
                        'website': 'https://www.goldsit.com',
                        'country': 'Turkey',
                        'models': {
                            'executive_desks': [
                                {'model': 'Goldsit Manager', 'price_range': '400-600', 'features': ['Melamine finish', 'Cable management', 'Storage included']},
                            ],
                            'workstations': [
                                {'model': 'Goldsit WorkHub', 'price_range': '250-400', 'features': ['Modular', 'Privacy screens', 'Cable tray']},
                            ],
                        }
                    },
                    {
                        'name': 'Kinwai',
                        'website': 'https://www.kinwai.com',
                        'country': 'China',
                        'models': {
                            'executive_desks': [
                                {'model': 'Kinwai Director Desk', 'price_range': '380-550', 'features': ['Veneer finish', 'Side return', 'Modesty panel']},
                            ],
                            'meeting_tables': [
                                {'model': 'Kinwai Conference', 'price_range': '300-500', 'features': ['Boat shape', 'Cable management', 'Wire tray']},
                            ],
                        }
                    },
                ]
            },
            'mid_range': {
                'seating': [
                    {
                        'name': 'Narbutas',
                        'website': 'https://www.narbutas.com',
                        'country': 'Lithuania',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Narbutas Navigo', 'price_range': '450-650', 'features': ['Premium mesh', 'Advanced ergonomics', 'Synchronized mechanism']},
                                {'model': 'Narbutas Vektor', 'price_range': '400-600', 'features': ['Leather upholstery', 'Full adjustability', 'Lumbar support']},
                            ],
                            'task_chairs': [
                                {'model': 'Narbutas Easy', 'price_range': '300-450', 'features': ['Mesh back', 'Ergonomic design', 'Multi-adjustment']},
                                {'model': 'Narbutas Wind', 'price_range': '280-420', 'features': ['Breathable mesh', 'Height adjustment', 'Tilt mechanism']},
                            ],
                            'conference_chairs': [
                                {'model': 'Narbutas Melody', 'price_range': '250-380', 'features': ['Upholstered', 'Stackable', 'Chrome frame']},
                            ],
                            'sofas': [
                                {'model': 'Narbutas Frame', 'price_range': '800-1200', 'features': ['Modular', 'Fabric options', 'Contemporary design']},
                            ],
                        }
                    },
                    {
                        'name': 'Frezza',
                        'website': 'https://www.frezza.com',
                        'country': 'Italy',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Frezza Director', 'price_range': '500-700', 'features': ['Italian design', 'Premium materials', 'Ergonomic excellence']},
                            ],
                            'task_chairs': [
                                {'model': 'Frezza Task', 'price_range': '320-480', 'features': ['Mesh back', 'Multi-function', 'Modern design']},
                            ],
                            'lounge_seating': [
                                {'model': 'Frezza Lounge', 'price_range': '600-900', 'features': ['Contemporary design', 'Comfortable', 'Various fabrics']},
                            ],
                        }
                    },
                    {
                        'name': 'Sedus',
                        'website': 'https://www.sedus.com',
                        'country': 'Germany',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Sedus Black Dot', 'price_range': '550-800', 'features': ['Award-winning design', 'Premium ergonomics', 'Net back']},
                                {'model': 'Sedus Swing Up', 'price_range': '480-680', 'features': ['Ergonomic', 'Synchronized mechanism', 'Premium materials']},
                            ],
                            'task_chairs': [
                                {'model': 'Sedus Se:flex', 'price_range': '350-520', 'features': ['Flexible back', 'Ergonomic', 'Multi-adjustment']},
                            ],
                        }
                    },
                    {
                        'name': 'Pedrali',
                        'website': 'https://www.pedrali.com',
                        'country': 'Italy',
                        'models': {
                            'visitor_chairs': [
                                {'model': 'Pedrali Nolita', 'price_range': '180-280', 'features': ['Modern design', 'Stackable', 'Multiple colors']},
                            ],
                            'conference_chairs': [
                                {'model': 'Pedrali Kira', 'price_range': '250-380', 'features': ['Contemporary', 'Upholstered', 'Chrome base']},
                            ],
                            'lounge_seating': [
                                {'model': 'Pedrali Buddy', 'price_range': '650-950', 'features': ['Lounge chair', 'Modern design', 'Comfortable']},
                            ],
                        }
                    },
                    {
                        'name': 'Sokoa',
                        'website': 'https://www.sokoa.fr',
                        'country': 'France',
                        'models': {
                            'task_chairs': [
                                {'model': 'Sokoa Arca', 'price_range': '320-470', 'features': ['Ergonomic', 'Mesh back', 'French design']},
                            ],
                            'conference_chairs': [
                                {'model': 'Sokoa Clea', 'price_range': '240-360', 'features': ['Stackable', 'Contemporary', 'Multiple finishes']},
                            ],
                        }
                    },
                    {
                        'name': 'Las Mobili',
                        'website': 'https://www.lasmobili.it',
                        'country': 'Italy',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Las Mobili Ego', 'price_range': '480-680', 'features': ['Italian craftsmanship', 'Ergonomic', 'Premium upholstery']},
                            ],
                            'task_chairs': [
                                {'model': 'Las Mobili Diva', 'price_range': '340-490', 'features': ['Mesh back', 'Ergonomic', 'Contemporary design']},
                            ],
                        }
                    },
                    {
                        'name': 'Nurus',
                        'website': 'https://www.nurus.com',
                        'country': 'Turkey',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Nurus IE', 'price_range': '520-720', 'features': ['Premium ergonomics', 'Turkish design', 'High quality']},
                            ],
                            'task_chairs': [
                                {'model': 'Nurus ME Too', 'price_range': '360-510', 'features': ['Mesh back', 'Multi-adjustment', 'Modern']},
                            ],
                        }
                    },
                    {
                        'name': 'B&T',
                        'website': 'https://www.bt-office.de',
                        'country': 'Germany',
                        'models': {
                            'executive_chairs': [
                                {'model': 'B&T Move', 'price_range': '490-690', 'features': ['Ergonomic excellence', 'German engineering', 'Premium materials']},
                            ],
                        }
                    },
                    {
                        'name': 'Forma 5',
                        'website': 'https://www.forma5.com',
                        'country': 'Spain',
                        'models': {
                            'task_chairs': [
                                {'model': 'Forma 5 One', 'price_range': '330-480', 'features': ['Spanish design', 'Ergonomic', 'Contemporary']},
                            ],
                            'conference_chairs': [
                                {'model': 'Forma 5 Bling', 'price_range': '260-390', 'features': ['Modern design', 'Stackable', 'Various finishes']},
                            ],
                        }
                    },
                ],
                'desking': [
                    {
                        'name': 'Narbutas',
                        'website': 'https://www.narbutas.com',
                        'country': 'Lithuania',
                        'models': {
                            'executive_desks': [
                                {'model': 'Narbutas Winea Pro', 'price_range': '800-1200', 'features': ['Premium finish', 'Cable management', 'European design']},
                            ],
                            'workstations': [
                                {'model': 'Narbutas Nova', 'price_range': '600-900', 'features': ['Modular system', 'Screen options', 'Cable management']},
                            ],
                            'meeting_tables': [
                                {'model': 'Narbutas E-place', 'price_range': '700-1000', 'features': ['Contemporary', 'Various sizes', 'Wire management']},
                            ],
                        }
                    },
                    {
                        'name': 'Frezza',
                        'website': 'https://www.frezza.com',
                        'country': 'Italy',
                        'models': {
                            'executive_desks': [
                                {'model': 'Frezza Pop', 'price_range': '900-1300', 'features': ['Italian design', 'Premium materials', 'Modern aesthetic']},
                            ],
                            'workstations': [
                                {'model': 'Frezza Time', 'price_range': '650-950', 'features': ['Modular', 'Contemporary', 'Flexible configuration']},
                            ],
                        }
                    },
                    {
                        'name': 'Sedus',
                        'website': 'https://www.sedus.com',
                        'country': 'Germany',
                        'models': {
                            'executive_desks': [
                                {'model': 'Sedus Temptation', 'price_range': '1000-1400', 'features': ['German engineering', 'Premium quality', 'Executive design']},
                            ],
                            'workstations': [
                                {'model': 'Sedus Mastermind', 'price_range': '700-1000', 'features': ['Flexible system', 'High quality', 'Ergonomic']},
                            ],
                        }
                    },
                    {
                        'name': 'Nurus',
                        'website': 'https://www.nurus.com',
                        'country': 'Turkey',
                        'models': {
                            'executive_desks': [
                                {'model': 'Nurus Flat', 'price_range': '850-1250', 'features': ['Contemporary', 'Cable management', 'Premium finish']},
                            ],
                            'workstations': [
                                {'model': 'Nurus IO', 'price_range': '680-980', 'features': ['Modular', 'Screens available', 'Modern design']},
                            ],
                        }
                    },
                ]
            },
            'high_end': {
                'seating': [
                    {
                        'name': 'Ofifran',
                        'website': 'https://www.ofifran.com',
                        'country': 'Spain',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Ofifran Executive Leather', 'price_range': '1200-1800', 'features': ['Genuine leather', 'Handcrafted', 'Premium ergonomics']},
                            ],
                            'sofas': [
                                {'model': 'Ofifran Lounge Collection', 'price_range': '2000-3500', 'features': ['Designer furniture', 'Premium materials', 'Customizable']},
                            ],
                        }
                    },
                    {
                        'name': 'Martex',
                        'website': 'https://www.martex.it',
                        'country': 'Italy',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Martex Status', 'price_range': '1400-2000', 'features': ['Italian luxury', 'Premium leather', 'Executive design']},
                            ],
                            'lounge_seating': [
                                {'model': 'Martex Soft', 'price_range': '1800-2800', 'features': ['Designer piece', 'Premium comfort', 'Italian craftsmanship']},
                            ],
                        }
                    },
                    {
                        'name': 'Uffix',
                        'website': 'https://www.uffix.com',
                        'country': 'Italy',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Uffix Nulite', 'price_range': '1100-1600', 'features': ['Contemporary Italian', 'Premium materials', 'Ergonomic excellence']},
                            ],
                            'task_chairs': [
                                {'model': 'Uffix Team', 'price_range': '800-1200', 'features': ['High-end task chair', 'Italian design', 'Multi-adjustment']},
                            ],
                        }
                    },
                    {
                        'name': 'Minotti',
                        'website': 'https://www.minotti.com',
                        'country': 'Italy',
                        'models': {
                            'sofas': [
                                {'model': 'Minotti Freeman', 'price_range': '5000-8000', 'features': ['Luxury design', 'Premium upholstery', 'Icon piece']},
                                {'model': 'Minotti Downs', 'price_range': '4500-7000', 'features': ['Designer sofa', 'Exceptional comfort', 'Italian luxury']},
                            ],
                            'lounge_seating': [
                                {'model': 'Minotti Bergere', 'price_range': '3000-5000', 'features': ['Classic design', 'Premium materials', 'Handcrafted']},
                            ],
                        }
                    },
                    {
                        'name': 'Marelli',
                        'website': 'https://www.marelli.it',
                        'country': 'Italy',
                        'models': {
                            'executive_chairs': [
                                {'model': 'Marelli Prestige', 'price_range': '1300-1900', 'features': ['Luxury executive', 'Premium leather', 'Italian design']},
                            ],
                        }
                    },
                    {
                        'name': 'Lacividina',
                        'website': 'https://www.lacividina.com',
                        'country': 'Italy',
                        'models': {
                            'sofas': [
                                {'model': 'Lacividina Olan', 'price_range': '3500-5500', 'features': ['Designer sofa', 'Modular options', 'Premium fabrics']},
                            ],
                            'lounge_seating': [
                                {'model': 'Lacividina Senso', 'price_range': '2500-4000', 'features': ['Contemporary design', 'Comfort focused', 'Italian quality']},
                            ],
                        }
                    },
                    {
                        'name': 'Tacchini',
                        'website': 'https://www.tacchini.it',
                        'country': 'Italy',
                        'models': {
                            'sofas': [
                                {'model': 'Tacchini Julep', 'price_range': '4000-6500', 'features': ['Award-winning design', 'Premium materials', 'Italian excellence']},
                            ],
                            'lounge_seating': [
                                {'model': 'Tacchini Sesann', 'price_range': '2800-4500', 'features': ['Designer chair', 'Exceptional comfort', 'Contemporary']},
                            ],
                        }
                    },
                    {
                        'name': 'B&B Italia',
                        'website': 'https://www.bebitalia.com',
                        'country': 'Italy',
                        'models': {
                            'sofas': [
                                {'model': 'B&B Italia Charles', 'price_range': '6000-10000', 'features': ['Iconic design', 'Luxury upholstery', 'Premium quality']},
                                {'model': 'B&B Italia Tufty-Time', 'price_range': '5500-9000', 'features': ['Modular luxury', 'Designer piece', 'Exceptional comfort']},
                            ],
                            'lounge_seating': [
                                {'model': 'B&B Italia Maxalto', 'price_range': '3500-6000', 'features': ['Executive lounge', 'Premium leather', 'Italian design']},
                            ],
                        }
                    },
                ],
                'desking': [
                    {
                        'name': 'Ofifran',
                        'website': 'https://www.ofifran.com',
                        'country': 'Spain',
                        'models': {
                            'executive_desks': [
                                {'model': 'Ofifran Executive Suite', 'price_range': '2500-4000', 'features': ['Premium veneer', 'Luxury finish', 'Cable management']},
                            ],
                        }
                    },
                    {
                        'name': 'Martex',
                        'website': 'https://www.martex.it',
                        'country': 'Italy',
                        'models': {
                            'executive_desks': [
                                {'model': 'Martex Presidente', 'price_range': '3000-4500', 'features': ['Italian luxury', 'Premium materials', 'Handcrafted']},
                            ],
                        }
                    },
                    {
                        'name': 'Uffix',
                        'website': 'https://www.uffix.com',
                        'country': 'Italy',
                        'models': {
                            'executive_desks': [
                                {'model': 'Uffix Unici', 'price_range': '2200-3500', 'features': ['Contemporary Italian', 'Premium finish', 'Modern design']},
                            ],
                            'meeting_tables': [
                                {'model': 'Uffix Level', 'price_range': '1800-3000', 'features': ['Conference table', 'Premium materials', 'Wire management']},
                            ],
                        }
                    },
                ]
            }
        }
    
    def get_brands_by_tier_and_category(self, tier: str, category: str) -> List[Dict]:
        """Get all brands for a specific tier and category"""
        tier_key = tier.lower().replace(' ', '_').replace('-', '_')
        category_key = category.lower()
        
        if tier_key in self.brands and category_key in self.brands[tier_key]:
            return self.brands[tier_key][category_key]
        return []
    
    def get_brand_models(self, tier: str, category: str, brand_name: str) -> Dict:
        """Get all models for a specific brand"""
        brands = self.get_brands_by_tier_and_category(tier, category)
        
        for brand in brands:
            if brand['name'].lower() == brand_name.lower():
                return brand.get('models', {})
        return {}
    
    def search_product(self, tier: str, category: str, subcategory: str, brand_name: Optional[str] = None) -> List[Dict]:
        """Search for products matching criteria"""
        results = []
        brands = self.get_brands_by_tier_and_category(tier, category)
        
        for brand in brands:
            if brand_name and brand['name'].lower() != brand_name.lower():
                continue
            
            models = brand.get('models', {})
            if subcategory in models:
                for model in models[subcategory]:
                    results.append({
                        'brand': brand['name'],
                        'country': brand['country'],
                        'website': brand['website'],
                        'model': model['model'],
                        'price_range': model['price_range'],
                        'features': model['features'],
                        'category': category,
                        'subcategory': subcategory,
                        'tier': tier
                    })
        
        return results
    
    def get_product(self, brand_name: str, category: str, subcategory: str, model_name: str, tier: str) -> Optional[Dict]:
        """
        Get a specific product by brand, category, subcategory, model, and tier
        
        Args:
            brand_name: Name of the brand
            category: Product category (e.g., 'seating', 'desking')
            subcategory: Product subcategory (e.g., 'executive_chairs', 'task_chairs')
            model_name: Product model name
            tier: Budget tier ('budgetary', 'mid_range', 'high_end')
        
        Returns:
            Product dictionary with description, image_url, etc., or None if not found
        """
        try:
            # Normalize tier name
            tier_key = tier.lower().replace(' ', '_').replace('-', '_')
            category_key = category.lower()
            
            # Get brands for this tier and category
            brands = self.get_brands_by_tier_and_category(tier_key, category_key)
            
            for brand in brands:
                if brand['name'].lower() != brand_name.lower():
                    continue
                
                models = brand.get('models', {})
                if subcategory in models:
                    for model in models[subcategory]:
                        if model.get('model', '').strip().lower() == model_name.strip().lower():
                            # Build product data
                            product = {
                                'brand': brand['name'],
                                'model': model.get('model', ''),
                                'description': ', '.join(model.get('features', [])) if model.get('features') else '',
                                'price_range': model.get('price_range', ''),
                                'features': model.get('features', []),
                                'category': category,
                                'subcategory': subcategory,
                                'tier': tier,
                                'country': brand.get('country', ''),
                                'website': brand.get('website', ''),
                                'image_url': None  # BrandDatabase doesn't store image URLs, use image_helper
                            }
                            
                            # Try to get image URL from image_helper
                            try:
                                from utils.image_helper import get_product_image_url
                                image_url = get_product_image_url(brand_name, category, subcategory, model_name, tier_key)
                                if image_url:
                                    product['image_url'] = image_url
                            except:
                                pass
                            
                            return product
            
            return None
        except Exception as e:
            logger.error(f"Error getting product: {e}")
            return None
    
    def get_all_tiers(self) -> List[str]:
        """Get all available budget tiers"""
        return ['Budgetary', 'Mid-Range', 'High-End']
    
    def get_all_categories(self) -> List[str]:
        """Get all available categories"""
        return ['Seating', 'Desking']
    
    def get_subcategories(self, category: str) -> List[str]:
        """Get subcategories for a specific category"""
        if category.lower() == 'seating':
            return ['executive_chairs', 'task_chairs', 'visitor_chairs', 'conference_chairs', 'sofas', 'lounge_seating']
        elif category.lower() == 'desking':
            return ['executive_desks', 'workstations', 'meeting_tables', 'pedestals', 'cabinets', 'lockers', 'partitions']
        return []


# Create a global instance for easy access
_db_instance = BrandDatabase()
BRAND_DATABASE = _db_instance.brands
