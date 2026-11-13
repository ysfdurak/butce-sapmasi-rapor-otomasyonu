// PQTest için boş veri kaynağı uzantısı: PQTest bir uzantı (.mez) ister; sorgular yalnız yerleşik işlevleri kullanır.
section Deneme;

[DataSource.Kind = "Deneme", Publish = "Deneme.Publish"]
shared Deneme.Contents = () => "Deneme";

Deneme = [Authentication = [Anonymous = []], Label = "Deneme"];
Deneme.Publish = [Beta = true, Category = "Other", ButtonText = {"Deneme", "Deneme"}];
