"""頁面本身與「有哪些線別／月份」的下拉資料。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來


@master_bp.route("/master")
def master_page():
    logo = ("logo_master.png"
            if os.path.exists(os.path.join(os.path.dirname(__file__), "static", "logo_master.png"))
            else "logo.png")
    return render_template("master.html", logo_file=logo, logged_in_user=_operator(),
                           is_admin=_is_admin(),
                           build_version=current_app.config.get("BUILD_VERSION", ""))


@master_bp.route("/api/master/lines")
def api_lines():
    cfg = _line_groups()
    conn = get_conn()
    try:
        raws = _raw_lines_seen(conn)
        months = [r["m"] for r in conn.execute(
            "SELECT DISTINCT substr(delivery_date, 1, 7) AS m FROM mst_orders "
            "WHERE delivery_date != '' ORDER BY m DESC").fetchall()]
    finally:
        conn.close()
    groups = sorted({_group_of(r, cfg) for r in raws})
    return jsonify({"groups": groups, "months": months, "this_month": _this_month()})

