"""Notion property <-> value mapping for the articles sync.

`PropertyMappingMixin` reads values out of Notion property payloads and builds
the payloads written to the target database. It relies on the `NotionArticlesSync`
state (`field_mapping`, `target_db`, `_schema_cache`) and `api_call`.
"""

from typing import Any, Dict



class PropertyMappingMixin:
    """Extract, format and map Notion properties between source and target."""

    def normalize_rowid(self, rowid: str) -> str:
        """Remove dashes from rowid for consistent format.
        
        Args:
            rowid: Row ID to normalize
            
        Returns:
            Normalized row ID without dashes
        """
        if rowid:
            return rowid.replace("-", "")
        return rowid
    
    def extract_value(self, prop: Dict[str, Any]) -> Any:
        """Extract value from a Notion property.
        
        Args:
            prop: Notion property dictionary
            
        Returns:
            Extracted value from the property
        """
        prop_type = prop.get("type")
        
        if prop_type == "title":
            return "".join([t.get("plain_text", "") for t in prop.get("title", [])])
        elif prop_type == "rich_text":
            return "".join([t.get("plain_text", "") for t in prop.get("rich_text", [])])
        elif prop_type == "select":
            select = prop.get("select")
            return select.get("name") if select else None
        elif prop_type == "multi_select":
            return [opt.get("name") for opt in prop.get("multi_select", [])]
        elif prop_type == "date":
            date_obj = prop.get("date")
            return date_obj.get("start") if date_obj else None
        elif prop_type == "checkbox":
            return prop.get("checkbox", False)
        elif prop_type == "url":
            return prop.get("url")
        elif prop_type == "number":
            return prop.get("number")
        elif prop_type == "created_time":
            return prop.get("created_time")
        elif prop_type == "last_edited_time":
            return prop.get("last_edited_time")
        elif prop_type == "formula":
            formula = prop.get("formula", {})
            formula_type = formula.get("type")
            if formula_type == "string":
                return formula.get("string")
            elif formula_type == "number":
                return formula.get("number")
            elif formula_type == "boolean":
                return formula.get("boolean")
            elif formula_type == "date":
                date_obj = formula.get("date")
                return date_obj.get("start") if date_obj else None
        elif prop_type == "relation":
            return [rel.get("id") for rel in prop.get("relation", [])]
        
        return None
    
    def format_property(self, field_name: str, value: Any) -> Dict[str, Any]:
        """Format a property value for Notion API.
        
        Args:
            field_name: Name of the field to format
            value: Value to format
            
        Returns:
            Formatted property dictionary for Notion API
        """
        # Get target schema
        if self.target_db not in self._schema_cache:
            result = self.api_call(f"databases/{self.target_db}")
            if result:
                self._schema_cache[self.target_db] = result.get("properties", {})
        
        schema = self._schema_cache.get(self.target_db, {})
        field_type = schema.get(field_name, {}).get("type", "rich_text")
        
        # Format based on type
        if field_type == "title":
            return {"title": [{"text": {"content": str(value)}}]}
        elif field_type == "rich_text":
            return {"rich_text": [{"text": {"content": str(value)}}]}
        elif field_type == "select":
            return {"select": {"name": str(value)}}
        elif field_type == "multi_select":
            if isinstance(value, list):
                return {"multi_select": [{"name": str(v)} for v in value]}
            return {"multi_select": [{"name": str(value)}]}
        elif field_type == "date":
            if isinstance(value, str):
                return {"date": {"start": value}}
            return {"date": {"start": str(value)}}
        elif field_type == "checkbox":
            return {"checkbox": bool(value)}
        elif field_type == "number":
            try:
                return {"number": float(value)}
            except (ValueError, TypeError):
                return {"number": 0}
        elif field_type == "url":
            return {"url": str(value)}
        else:
            return {"rich_text": [{"text": {"content": str(value)}}]}
    
    def map_properties(self, source_props: Dict[str, Any]) -> Dict[str, Any]:
        """Map properties from source to target format.
        
        Args:
            source_props: Source properties dictionary
            
        Returns:
            Mapped properties dictionary for target format
        """
        mapped = {}
        
        for source_field, target_field in self.field_mapping.items():
            if source_field in source_props:
                value = self.extract_value(source_props[source_field])
                
                # Normalize rowid values
                if source_field == "rowid" and value:
                    value = self.normalize_rowid(value)
                
                if value is not None:
                    mapped[target_field] = self.format_property(target_field, value)
        
        return mapped
