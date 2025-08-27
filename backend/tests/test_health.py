def test_healthz_ok(client):
    res = client.get("/healthz")
    assert res.status_code == 200
    data = res.get_json()
    assert data["ok"] is True
    assert data["db"] == "up"
