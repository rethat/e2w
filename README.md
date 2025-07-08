# E2W project
Export template HTML tag to Word

Sample to use
```
    context = {
            "tag_name1": "Value name1",
            "tag_name2": "Customer X",
            "api_headers": { "Content-Type":"application/json", "Authorization":"Bearer ***" },        
            "apis": {
                "https://api.example.com/products": {
                    "params": { "id": 123, "code": "PR0D01" } 
                    },
                "https://api.example.com/summary": {
                    "params": {"order_id": 10, "token": "xyAf0xdDAv23" }
                    }
            }
        }

    e2w = ExportToWord(
        context=context,
        template="templates/sample.template",
        output_path= f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-output.docx"
    )
    e2w.render()
```

