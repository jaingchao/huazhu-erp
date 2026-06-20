<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>华住进销存ERP · 第29版 · 同比异常+搜索跳转</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js">
    </script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js">
    </script>
    <style>
        :root {
            --bg: #1a1f2e;
            --sidebar-bg: #141824;
            --card: rgba(28, 34, 48, 0.95);
            --gold: #c8a84e;
            --gold2: #f0d878;
            --red: #d4554a;
            --green: #4a9e6e;
            --blue: #4a7fb5;
            --border: rgba(200, 168, 78, 0.15);
            --text: #c0c8d4;
            --text2: #8899aa;
            --shadow: 0 2px 12px rgba(0, 0, 0, 0.3)
        }
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
            display: flex
        }
        .sidebar {
            width: 220px;
            min-width: 220px;
            background: var(--sidebar-bg);
            padding: 20px 0;
            display: flex;
            flex-direction: column;
            border-right: 1px solid var(--border);
            position: sticky;
            top: 0;
            height: 100vh;
            z-index: 10
        }
        .sidebar-logo {
            padding: 0 20px 20px;
            border-bottom: 1px solid var(--border);
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 10px
        }
        .sidebar-logo .icon {
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, var(--gold), #a08030);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.1em
        }
        .sidebar-logo .title {
            font-weight: 700;
            font-size: .95em;
            color: var(--gold2);
            letter-spacing: 1px
        }
        .sidebar-menu {
            flex: 1;
            padding: 0 12px
        }
        .sidebar-menu .menu-item {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 14px;
            border-radius: 8px;
            cursor: pointer;
            font-size: .85em;
            font-weight: 500;
            color: var(--text2);
            transition: all .2s;
            margin-bottom: 2px;
            border: 1px solid transparent
        }
        .sidebar-menu .menu-item:hover {
            background: rgba(200, 168, 78, 0.06);
            color: var(--text);
            border-color: rgba(200, 168, 78, 0.15)
        }
        .sidebar-menu .menu-item.active {
            background: rgba(200, 168, 78, 0.12);
            color: var(--gold);
            border-color: rgba(200, 168, 78, 0.3);
            font-weight: 600
        }
        .sidebar-footer {
            padding: 12px 20px;
            border-top: 1px solid var(--border);
            font-size: .7em;
            color: var(--text2);
            text-align: center
        }
        .main-content {
            flex: 1;
            padding: 20px 24px;
            overflow-y: auto;
            max-height: 100vh
        }
        .page {
            display: none
        }
        .page.active {
            display: block
        }
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 10px
        }
        .page-header h2 {
            font-size: 1.1em;
            font-weight: 700;
            color: var(--gold2);
            letter-spacing: 1px
        }
        .month-badge {
            background: rgba(200, 168, 78, 0.1);
            border: 1px solid rgba(200, 168, 78, 0.3);
            border-radius: 16px;
            padding: 6px 14px;
            color: var(--gold);
            font-weight: 600;
            font-size: .8em;
            cursor: pointer;
            transition: all .2s;
            display: inline-flex;
            align-items: center;
            gap: 8px
        }
        .month-badge:hover {
            background: rgba(200, 168, 78, 0.18);
            border-color: var(--gold2)
        }
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 8px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            cursor: pointer;
            font-size: .78em;
            font-weight: 600;
            transition: all .2s;
            background: rgba(255, 255, 255, 0.03);
            color: var(--text)
        }
        .btn:hover {
            border-color: var(--gold);
            transform: translateY(-1px);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2)
        }
        .btn-primary {
            background: linear-gradient(135deg, rgba(200, 168, 78, 0.25), rgba(160, 128, 48, 0.2));
            border-color: var(--gold);
            color: var(--gold2);
            font-weight: 700
        }
        .btn-primary:hover {
            background: linear-gradient(135deg, rgba(200, 168, 78, 0.35), rgba(160, 128, 48, 0.3));
            box-shadow: 0 2px 16px rgba(200, 168, 78, 0.2)
        }
        .btn-green {
            border-color: rgba(74, 158, 110, 0.5);
            color: var(--green)
        }
        .btn-blue {
            border-color: rgba(74, 127, 181, 0.5);
            color: #7ab8f0
        }
        .btn-red {
            border-color: rgba(212, 85, 74, 0.5);
            color: #ff6b6b
        }
        .btn-sm {
            padding: 5px 10px;
            font-size: .72em;
            border-radius: 6px
        }
        .dashboard-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 12px;
            margin-bottom: 20px
        }
        .stat-card {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 18px 16px;
            text-align: center;
            box-shadow: var(--shadow);
            transition: all .3s
        }
        .stat-card:hover {
            border-color: var(--gold);
            transform: translateY(-2px);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4)
        }
        .stat-card .icon {
            font-size: 1.6em;
            margin-bottom: 6px
        }
        .stat-card .num {
            font-size: 1.8em;
            font-weight: 800;
            background: linear-gradient(135deg, var(--gold2), var(--gold));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent
        }
        .stat-card .num.warn {
            background: linear-gradient(135deg, #ff6b6b, var(--red));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent
        }
        .stat-card .num.green {
            background: linear-gradient(135deg, #4a9e6e, #2d8a5e);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent
        }
        .stat-card .num.yellow {
            background: linear-gradient(135deg, #f0d878, #c8a84e);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent
        }
        .stat-card .num.red {
            background: linear-gradient(135deg, #ff6b6b, #d4554a);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent
        }
        .stat-card .label {
            font-size: .7em;
            color: var(--text2);
            margin-top: 4px;
            letter-spacing: 1px
        }
        .stat-card .sub {
            font-size: .6em;
            color: var(--text2);
            margin-top: 2px
        }
        .alert-banner {
            background: rgba(212, 85, 74, 0.08);
            border: 1px solid rgba(212, 85, 74, 0.2);
            border-radius: 10px;
            padding: 10px 16px;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 12px;
            flex-wrap: wrap
        }
        .alert-banner .label {
            font-weight: 700;
            color: #ff6b6b;
            font-size: .85em
        }
        .alert-banner .detail {
            color: var(--text2);
            font-size: .8em;
            flex: 1
        }
        .alert-banner .detail strong {
            color: var(--text)
        }
        .alert-banner .btn-sm {
            font-size: .7em
        }
        .alert-banner.gold {
            background: rgba(200, 168, 78, 0.08);
            border-color: rgba(200, 168, 78, 0.25)
        }
        .alert-banner.gold .label {
            color: var(--gold2)
        }
        .purchase-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(210px, 1fr));
            gap: 8px;
            margin-bottom: 16px;
            max-height: 260px;
            overflow-y: auto;
            padding: 2px 0
        }
        .purchase-item {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 10px 12px;
            transition: all .2s
        }
        .purchase-item:hover {
            border-color: var(--gold)
        }
        .purchase-item .pi-name {
            font-weight: 600;
            color: var(--gold2);
            font-size: .8em
        }
        .purchase-item .pi-detail {
            font-size: .7em;
            color: var(--text2);
            display: flex;
            justify-content: space-between;
            margin-top: 3px
        }
        .purchase-item .pi-detail .val {
            color: var(--text);
            font-weight: 600
        }
        .purchase-item .pi-urgent {
            color: #ff6b6b;
            font-weight: 700
        }
        .purchase-item .pi-normal {
            color: var(--gold)
        }
        .chart-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
            gap: 14px;
            margin-bottom: 20px
        }
        .chart-box {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 16px;
            box-shadow: var(--shadow)
        }
        .chart-box h3 {
            color: var(--gold);
            font-size: .85em;
            margin-bottom: 10px;
            font-weight: 600
        }
        .chart-wrap {
            position: relative;
            width: 100%;
            height: 280px
        }
        .product-select {
            padding: 6px 12px;
            background: var(--sidebar-bg);
            border: 1px solid var(--border);
            border-radius: 6px;
            color: var(--gold);
            font-size: .82em;
            font-weight: 600;
            cursor: pointer;
            margin-bottom: 10px;
            max-width: 250px
        }
        .product-select:focus {
            outline: none;
            border-color: var(--gold)
        }
        .view-tabs {
            display: flex;
            gap: 6px;
            margin-bottom: 14px;
            flex-wrap: wrap
        }
        .view-tab {
            padding: 7px 16px;
            border: 1px solid var(--border);
            border-radius: 20px;
            cursor: pointer;
            font-size: .78em;
            font-weight: 600;
            color: var(--text2);
            transition: all .2s;
            background: transparent
        }
        .view-tab.active {
            background: rgba(200, 168, 78, 0.15);
            border-color: var(--gold);
            color: var(--gold)
        }
        .table-wrap {
            overflow-x: auto;
            border-radius: 10px;
            border: 1px solid var(--border);
            max-height: 50vh;
            overflow-y: auto;
            background: var(--card);
            box-shadow: var(--shadow);
            margin-bottom: 20px
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: .74em;
            min-width: 900px
        }
        th {
            background: var(--sidebar-bg);
            color: var(--gold);
            padding: 9px 6px;
            text-align: center;
            font-weight: 700;
            white-space: nowrap;
            border-bottom: 2px solid rgba(200, 168, 78, 0.3);
            position: sticky;
            top: 0;
            z-index: 2;
            cursor: pointer
        }
        th:hover {
            background: #1e2640
        }
        td {
            padding: 5px 4px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.03);
            white-space: nowrap;
            text-align: center
        }
        tr:hover td {
            background: rgba(200, 168, 78, 0.03)
        }
        tr.warn-row td {
            background: rgba(212, 85, 74, 0.05)
        }
        tr.highlight td {
            background: rgba(200, 168, 78, 0.15) !important;
            border-color: var(--gold)
        }
        .cell-input {
            width: 46px;
            padding: 4px 2px;
            background: transparent;
            border: 1px solid transparent;
            border-radius: 3px;
            color: #fff;
            text-align: center;
            font-size: .82em;
            transition: all .15s
        }
        .cell-input:hover {
            border-color: rgba(200, 168, 78, 0.3);
            background: rgba(200, 168, 78, 0.05)
        }
        .cell-input:focus {
            border-color: var(--gold);
            background: #1a2030;
            outline: none;
            box-shadow: 0 0 0 2px rgba(200, 168, 78, 0.15)
        }
        .cell-input-safe {
            width: 52px;
            padding: 4px 2px;
            background: rgba(200, 168, 78, 0.05);
            border: 1px solid rgba(200, 168, 78, 0.25);
            border-radius: 3px;
            color: var(--gold);
            text-align: center;
            font-size: .82em;
            transition: all .15s;
            font-weight: 600
        }
        .cell-input-safe:hover {
            border-color: var(--gold);
            background: rgba(200, 168, 78, 0.12)
        }
        .cell-input-safe:focus {
            border-color: var(--gold);
            background: #1a2030;
            outline: none;
            box-shadow: 0 0 0 2px rgba(200, 168, 78, 0.2)
        }
        .cell-sum {
            font-weight: 700;
            color: var(--gold)
        }
        .in-sum {
            color: #3fb950
        }
        .out-sum {
            color: #e94560
        }
        .tag {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 10px;
            font-size: .7em;
            font-weight: 700
        }
        .tag-danger {
            background: rgba(212, 85, 74, 0.2);
            color: #ff6b6b
        }
        .tag-warn {
            background: rgba(200, 168, 78, 0.2);
            color: var(--gold2)
        }
        .tag-ok {
            background: rgba(74, 158, 110, 0.2);
            color: var(--green)
        }
        .tag-info {
            background: rgba(74, 127, 181, 0.2);
            color: #7ab8f0
        }
        .stock-cards-wall {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(195px, 1fr));
            gap: 10px
        }
        .stock-card {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 14px 16px;
            transition: all .3s;
            position: relative;
            box-shadow: var(--shadow)
        }
        .stock-card:hover {
            border-color: var(--gold);
            transform: translateY(-2px);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4)
        }
        .stock-card .sc-name {
            font-weight: 700;
            color: var(--gold2);
            font-size: .88em;
            margin-bottom: 4px
        }
        .stock-card .sc-spec {
            font-size: .65em;
            color: var(--text2);
            margin-bottom: 8px
        }
        .stock-card .sc-stock {
            font-size: 1.7em;
            font-weight: 800;
            margin-bottom: 4px
        }
        .stock-card .sc-info {
            display: flex;
            gap: 12px;
            font-size: .68em;
            color: var(--text2);
            margin-bottom: 8px
        }
        .stock-card .sc-info span {
            color: var(--text);
            font-weight: 600
        }
        .stock-card .sc-safe {
            display: flex;
            align-items: center;
            gap: 4px;
            font-size: .68em;
            color: var(--text2)
        }
        .stock-card .sc-safe input {
            width: 44px;
            padding: 2px 4px;
            background: rgba(200, 168, 78, 0.06);
            border: 1px solid rgba(200, 168, 78, 0.25);
            border-radius: 3px;
            color: var(--gold);
            text-align: center;
            font-size: .9em;
            font-weight: 600
        }
        .stock-card .sc-safe input:focus {
            border-color: var(--gold);
            outline: none;
            background: #1a2030
        }
        .stock-card .sc-status {
            position: absolute;
            top: 10px;
            right: 10px;
            font-size: .6em;
            padding: 3px 10px;
            border-radius: 10px;
            font-weight: 700
        }
        .status-ok {
            background: rgba(74, 158, 110, 0.15);
            color: var(--green)
        }
        .status-warn {
            background: rgba(200, 168, 78, 0.15);
            color: var(--gold)
        }
        .status-danger {
            background: rgba(212, 85, 74, 0.15);
            color: #ff6b6b
        }
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            z-index: 100;
            display: flex;
            justify-content: center;
            align-items: center
        }
        .modal-box {
            background: #1a2030;
            border: 1px solid var(--gold);
            border-radius: 14px;
            padding: 24px;
            width: 90%;
            max-width: 580px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
            max-height: 80vh;
            overflow-y: auto
        }
        .modal-box h3 {
            color: var(--gold);
            margin-bottom: 16px;
            font-size: 1em
        }
        .modal-box input {
            width: 100%;
            padding: 9px 12px;
            background: #0f1520;
            border: 1px solid #2a3040;
            border-radius: 6px;
            color: #fff;
            font-size: .85em;
            margin-bottom: 8px
        }
        .modal-box input:focus {
            border-color: var(--gold);
            outline: none
        }
        .modal-box .preview-table {
            font-size: .75em;
            max-height: 200px;
            overflow-y: auto;
            margin: 12px 0;
            border: 1px solid var(--border);
            border-radius: 6px
        }
        .modal-box .preview-table table {
            min-width: unset;
            font-size: .7em
        }
        .modal-box .preview-table th,
        .modal-box .preview-table td {
            padding: 4px 8px
        }
        .history-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 16px
        }
        .history-chip {
            padding: 6px 14px;
            border: 1px solid var(--border);
            border-radius: 16px;
            cursor: pointer;
            font-size: .8em;
            color: var(--text2);
            transition: all .2s
        }
        .history-chip:hover,
        .history-chip.active {
            border-color: var(--gold);
            color: var(--gold);
            background: rgba(200, 168, 78, 0.1)
        }
        .quick-input-wrap {
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 16px;
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 12px 16px
        }
        .quick-input-wrap input {
            flex: 1;
            min-width: 200px;
            padding: 8px 14px;
            background: #0f1520;
            border: 1px solid #2a3040;
            border-radius: 6px;
            color: #fff;
            font-size: .82em
        }
        .quick-input-wrap input:focus {
            border-color: var(--gold);
            outline: none
        }
        .quick-input-wrap .hint {
            font-size: .7em;
            color: var(--text2);
            white-space: nowrap
        }
        .quick-result {
            font-size: .75em;
            color: var(--green);
            padding: 4px 0
        }
        .btn-today {
            background: rgba(74, 158, 110, 0.15);
            border: 1px solid rgba(74, 158, 110, 0.3);
            color: #3fb950;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: .7em;
            cursor: pointer;
            transition: all .2s;
            font-weight: 600
        }
        .btn-today:hover {
            background: rgba(74, 158, 110, 0.3);
            border-color: #3fb950
        }
        .btn-today-out {
            background: rgba(233, 69, 96, 0.15);
            border: 1px solid rgba(233, 69, 96, 0.3);
            color: #e94560
        }
        .btn-today-out:hover {
            background: rgba(233, 69, 96, 0.3);
            border-color: #e94560
        }
        .month-nav {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap
        }
        .summary-text {
            background: rgba(200, 168, 78, 0.04);
            border-left: 3px solid var(--gold);
            padding: 12px 16px;
            border-radius: 6px;
            font-size: .8em;
            line-height: 1.8;
            color: var(--text);
            margin: 10px 0;
            white-space: pre-wrap;
            max-height: 200px;
            overflow-y: auto
        }
        .summary-text strong {
            color: var(--gold2)
        }
        .search-wrap {
            display: flex;
            align-items: center;
            gap: 6px;
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border);
            border-radius: 20px;
            padding: 4px 12px 4px 16px;
            transition: all .2s
        }
        .search-wrap:focus-within {
            border-color: var(--gold);
            background: rgba(255, 255, 255, 0.05)
        }
        .search-wrap input {
            background: transparent;
            border: none;
            padding: 6px 0;
            color: var(--text);
            font-size: .8em;
            width: 160px;
            outline: none
        }
        .search-wrap input::placeholder {
            color: var(--text2)
        }
        .search-wrap .icon {
            color: var(--text2);
            font-size: .8em
        }
        .search-results {
            position: absolute;
            top: 100%;
            right: 0;
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 10px;
            min-width: 220px;
            max-height: 300px;
            overflow-y: auto;
            z-index: 20;
            box-shadow: var(--shadow);
            display: none;
            margin-top: 4px
        }
        .search-results .item {
            padding: 8px 14px;
            cursor: pointer;
            font-size: .78em;
            border-bottom: 1px solid rgba(255, 255, 255, 0.03);
            transition: all .15s
        }
        .search-results .item:hover {
            background: rgba(200, 168, 78, 0.08);
            color: var(--gold)
        }
        .search-results .item .code {
            color: var(--text2);
            font-size: .7em
        }
        .search-wrap {
            position: relative
        }
        .search-wrap .clear-btn {
            color: var(--text2);
            cursor: pointer;
            font-size: .7em;
            padding: 2px 6px;
            display: none
        }
        .search-wrap .clear-btn:hover {
            color: var(--text)
        }
        @media(max-width:768px) {
            .sidebar {
                width: 60px;
                min-width: 60px
            }
            .sidebar-logo .title,
            .sidebar-menu .menu-item span,
            .sidebar-footer {
                display: none
            }
            .sidebar-menu .menu-item {
                justify-content: center;
                padding: 10px 0
            }
            .main-content {
                padding: 12px
            }
            .chart-grid {
                grid-template-columns: 1fr
            }
            .stock-cards-wall {
                grid-template-columns: repeat(2, 1fr)
            }
            .quick-input-wrap {
                flex-direction: column;
                align-items: stretch
            }
            .quick-input-wrap .hint {
                white-space: normal
            }
            .purchase-grid {
                grid-template-columns: 1fr
            }
            .search-wrap input {
                width: 100px
            }
            .search-results {
                right: auto;
                left: 0;
                min-width: 180px
            }
        }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-logo"><div class="icon">📦</div><div class="title">华住进销存ERP</div></div>
        <div class="sidebar-menu">
            <div class="menu-item active" onclick="switchPage('dashboard')" data-page="dashboard">📊 <span>仪表盘</span></div>
            <div class="menu-item" onclick="switchPage('inpage')" data-page="inpage">📥 <span>入库管理</span></div>
            <div class="menu-item" onclick="switchPage('outpage')" data-page="outpage">📤 <span>出库管理</span></div>
            <div class="menu-item" onclick="switchPage('stockpage')" data-page="stockpage">📋 <span>库存总览</span></div>
            <div class="menu-item" onclick="switchPage('chartpage')" data-page="chartpage">📈 <span>数据分析</span></div>
            <div class="menu-item" onclick="switchPage('historypage')" data-page="historypage">📅 <span>历史记录</span></div>
        </div>
        <div class="sidebar-footer">v29 · 同比异常+搜索<br>数据存于本地浏览器</div>
    </div>
    <div class="main-content">
        <div class="page active" id="page-dashboard">
            <div class="page-header"><h2>📊 数据仪表盘</h2>
                <div class="month-nav">
                    <span class="month-badge" id="monthBadge">📅 2026年6月</span>
                    <button class="btn btn-sm btn-primary" onclick="checkMonthRollover()">🔄 检查月份</button>
                    <div class="search-wrap" id="searchWrap">
                        <span class="icon">🔍</span>
                        <input id="searchInput" placeholder="搜索产品..." oninput="handleSearch(this.value)" onfocus="handleSearch(this.value)">
                        <span class="clear-btn" id="searchClear" onclick="clearSearch()">✕</span>
                        <div class="search-results" id="searchResults"></div>
                    </div>
                </div>
            </div>
            <div id="alertBannerContainer"></div>
            <div id="yoyAlertContainer"></div>
            <div id="purchaseContainer"></div>
            <div class="quick-input-wrap">
                <input id="quickInput" placeholder="例如：全季6.5盎司 入库 50  或  汉庭9盎司 出库 30" onkeydown="if(event.key==='Enter')quickBatch()">
                <button class="btn btn-primary btn-sm" onclick="quickBatch()">⏎ 执行</button>
                <span class="hint">支持：产品名/规格/编码 + 入库/出库 + 数量</span>
            </div>
            <div id="quickResult" class="quick-result"></div>
            <div class="dashboard-cards" id="stats"></div>
            <div class="chart-grid">
                <div class="chart-box"><h3>📈 出入库趋势（1-31号）</h3><div class="chart-wrap"><canvas id="dailyChart"></canvas></div></div>
                <div class="chart-box"><h3>🔔 库存预警占比</h3><div class="chart-wrap" style="max-width:350px;margin:0 auto"><canvas id="alertChart"></canvas></div></div>
            </div>
            <div class="chart-box" style="margin-top:14px"><h3>📊 全部产品月出入库柱状图</h3><div class="chart-wrap"><canvas id="barAllChart" style="max-height:450px"></canvas></div></div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:10px">
                <button class="btn btn-primary" onclick="openAddProduct()">➕ 新增产品</button>
                <button class="btn btn-primary" onclick="monthlyCarryover()">🔄 月度结转</button>
                <button class="btn btn-green" onclick="exportExcel()">📤 导出总表</button>
                <button class="btn btn-blue" onclick="document.getElementById('importJsonFile').click()">📥 导入JSON</button>
                <button class="btn" onclick="exportData()">📤 导出JSON</button>
                <button class="btn btn-blue" onclick="document.getElementById('importInFile').click()">📥 导入入库Excel</button>
                <button class="btn btn-blue" onclick="document.getElementById('importOutFile').click()">📥 导入出库Excel</button>
            </div>
        </div>
        <div class="page" id="page-inpage">
            <div class="page-header"><h2>📥 入库管理</h2><div style="display:flex;gap:6px;flex-wrap:wrap"><button class="btn btn-green btn-sm" onclick="exportInExcel()">📤 导出入库表</button></div></div>
            <div class="table-wrap" style="max-height:65vh"><table><thead><tr><th>产品</th><th>规格</th><th>编码</th><th>盘点</th><th>安全线</th><script>
                        for (let i = 1; i <= 31; i++) document.write(`<th>${i}日</th>`);
                    </script><th>月汇总</th><th>操作</th></tr></thead><tbody id="inTbody"></tbody></table></div>
        </div>
        <div class="page" id="page-outpage">
            <div class="page-header"><h2>📤 出库管理</h2><div style="display:flex;gap:6px;flex-wrap:wrap"><button class="btn btn-green btn-sm" onclick="exportOutExcel()">📤 导出出库表</button></div></div>
            <div class="table-wrap" style="max-height:65vh"><table><thead><tr><th>产品</th><th>规格</th><th>编码</th><script>
                        for (let i = 1; i <= 31; i++) document.write(`<th>${i}日</th>`);
                    </script><th>月汇总</th><th>发货天</th><th>日均</th><th>实时库存</th><th>可发天</th><th>状态</th><th>操作</th></tr></thead><tbody id="outTbody"></tbody></table></div>
        </div>
        <div class="page" id="page-stockpage">
            <div class="page-header"><h2>📋 库存总览</h2><span style="font-size:.78em;color:var(--text2)">按库存从少到多排列 | 点击安全库存值可直接编辑</span></div>
            <div class="view-tabs">
                <span class="view-tab active" id="tabStockCards" onclick="switchStockView('cards')">🃏 卡片格式</span>
                <span class="view-tab" id="tabStockTable" onclick="switchStockView('table')">📋 表格格式</span>
            </div>
            <div id="stockCardsView"><div class="stock-cards-wall" id="stockCardsWall"></div></div>
            <div id="stockTableView" class="hidden"><div class="table-wrap" style="max-height:65vh"><table><thead><tr><th>产品</th><th>规格</th><th>编码</th><th>盘点</th><th>实时库存</th><th>安全库存</th><th>可发天</th><th>日均</th><th>状态</th></tr></thead><tbody id="stockTbody"></tbody></table></div></div>
        </div>
        <div class="page" id="page-chartpage">
            <div class="page-header"><h2>📈 数据分析</h2></div>
            <div class="chart-grid">
                <div class="chart-box"><h3>📈 单品出入库走势</h3><select class="product-select" id="singleProductSelect" onchange="updateSingleChart()"></select><div class="chart-wrap"><canvas id="singleChart"></canvas></div></div>
                <div class="chart-box"><h3>📊 单品每日出入库柱状图</h3><div class="chart-wrap"><canvas id="singleBarChart"></canvas></div></div>
            </div>
        </div>
        <div class="page" id="page-historypage">
            <div class="page-header"><h2>📅 历史记录</h2></div>
            <div class="history-list" id="historyList"></div>
            <div id="historySummaryContainer"></div>
            <div class="table-wrap" style="max-height:55vh"><table><thead><tr><th>产品</th><th>规格</th><th>编码</th><th>盘点</th><th>安全线</th><script>
                        for (let i = 1; i <= 31; i++) document.write(`<th>${i}日</th>`);
                    </script><th>月入库</th><th>月出库</th><th>发货天</th><th>日均</th><th>实时库存</th><th>可发天</th></tr></thead><tbody id="historyTbody"></tbody></table></div>
        </div>
    </div>

    <div class="modal-overlay" id="addProductModal" style="display:none"><div class="modal-box"><h3 id="modalTitle">➕ 新增产品</h3><input id="mName" placeholder="产品名称"><input id="mSpec" placeholder="产品规格"><input id="mCode" placeholder="产品编码"><input id="mSafeStock" type="number" placeholder="安全库存预警线" value="10"><input id="mInitStock" type="number" placeholder="盘点库存" value="0"><input type="hidden" id="mEditId"><div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px"><button class="btn" onclick="closeAddProduct()">取消</button><button class="btn btn-primary" onclick="saveProduct()">💾 保存</button></div></div></div>
    <div class="modal-overlay" id="carryoverModal" style="display:none"><div class="modal-box"><h3>🔄 月度结转</h3><p id="carryoverMsg" style="color:#8899aa;margin-bottom:12px;font-size:.85em"></p><div id="carryoverSummary" class="summary-text" style="display:none"></div><div style="display:flex;gap:8px;justify-content:center;margin-top:16px"><button class="btn" onclick="document.getElementById('carryoverModal').style.display='none'">取消</button><button class="btn btn-primary" onclick="confirmCarryover()">✅ 确认结转</button></div></div></div>
    <div class="modal-overlay" id="importPreviewModal" style="display:none"><div class="modal-box"><h3>📥 导入预览</h3><div id="importPreviewContent"></div><div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px"><button class="btn" onclick="document.getElementById('importPreviewModal').style.display='none'">取消</button><button class="btn btn-primary" id="importConfirmBtn" onclick="confirmImport()">✅ 确认导入</button></div></div></div>
    <div class="modal-overlay" id="alertDetailModal" style="display:none"><div class="modal-box"><h3>⚠️ 异常波动详情</h3><div id="alertDetailContent"></div><div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px"><button class="btn" onclick="document.getElementById('alertDetailModal').style.display='none'">关闭</button></div></div></div>
    <div class="modal-overlay" id="yoyDetailModal" style="display:none"><div class="modal-box"><h3>📊 月度同比异常详情</h3><div id="yoyDetailContent"></div><div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px"><button class="btn" onclick="document.getElementById('yoyDetailModal').style.display='none'">关闭</button></div></div></div>
    <input type="file" id="importJsonFile" accept=".json" style="display:none" onchange="handleJsonImport(event)">
    <input type="file" id="importInFile" accept=".xlsx,.xls" style="display:none" onchange="handleInExcelImport(event)">
    <input type="file" id="importOutFile" accept=".xlsx,.xls" style="display:none" onchange="handleOutExcelImport(event)">

    <script>
        const STORAGE_KEY = 'erp_pro_v5';
        let data = { products: [], currentYear: 2026, currentMonth: 6, history: {} };
        let dailyChart, singleChart, singleBarChart, barAllChart, alertChart;
        let currentPage = 'dashboard';
        let stockView = 'cards';
        let pendingImportData = null;
        let detectedAlerts = [];
        let yoyAlerts = [];
        let searchTimeout = null;
        const PRODUCT_BASE = [{ name: '汉庭酒店', spec: '9盎司', code: '849479', safeStock: 21 },
            { name: '全季酒店', spec: '6.5盎司', code: '816545', safeStock: 28 },
            { name: '全季(晓山青)', spec: '16盎司', code: '1123064', safeStock: 5 },
            { name: '全季(晓山青)盖子', spec: '90口径', code: '1123065', safeStock: 4 },
            { name: '全季酒店', spec: '8盎司中空', code: '1170947', safeStock: 2 },
            { name: '全季酒店盖子', spec: '80口径', code: '1170948', safeStock: 2 },
            { name: '桔子酒店', spec: '6.5盎司', code: '955902', safeStock: 11 },
            { name: '桔子酒店', spec: '8盎司中空', code: '955901', safeStock: 4 },
            { name: '桔子水晶酒店', spec: '6.5盎司', code: '816541', safeStock: 5 },
            { name: '美仑酒店', spec: '9盎司', code: '842111', safeStock: 7 },
            { name: '美仑国际酒店', spec: '9盎司', code: '842112', safeStock: 5 },
            { name: '美仑美奂酒店', spec: '9盎司', code: '850078', safeStock: 1 },
            { name: '禧玥（黑）', spec: '14盎司中空', code: '816644', safeStock: 1 },
            { name: '禧玥（白）', spec: '14盎司中空', code: '816645', safeStock: 1 },
            { name: '禧玥', spec: '8盎司中空', code: '816534', safeStock: 1 },
            { name: '禧玥', spec: '6.5盎司', code: '816546', safeStock: 1 },
            { name: '宜必思酒店', spec: '6.5盎司', code: '1052152', safeStock: 2 },
            { name: '宜必思酒店', spec: '9盎司', code: '816550', safeStock: 7 },
            { name: '宜必思尚品酒店', spec: '6.5盎司', code: '816552', safeStock: 2 },
            { name: '宜必思尚品酒店', spec: '9盎司', code: '870858', safeStock: 10 },
            { name: '星程酒店', spec: '6.5盎司', code: '816547', safeStock: 7 },
            { name: '美居酒店', spec: '6.5盎司', code: '816544', safeStock: 6 },
            { name: '漫心酒店', spec: '6.5盎司', code: '847833', safeStock: 5 },
            { name: '城家公寓酒店', spec: '6.5盎司', code: '816535', safeStock: 1 },
            { name: '你好酒店', spec: '9盎司', code: '816525', safeStock: 6 },
            { name: '你好酒店黄绿（独立包装）', spec: '9盎司', code: '822324', safeStock: 8 },
            { name: '海友酒店（独立包装）', spec: '9盎司', code: '1088277', safeStock: 6 },
            { name: '怡莱酒店（独立包装）', spec: '9盎司', code: '822325', safeStock: 1 },
            { name: '欢阁酒店（CitiG0）', spec: '6.5盎司', code: '820178', safeStock: 1 },
            { name: '城际酒店（Intercity）', spec: '9盎司', code: '816524', safeStock: 3 },
            { name: '黑色盖子', spec: '90口径', code: '816560', safeStock: 1 },
            { name: '黑色盖子', spec: '80口径', code: '956922', safeStock: 3 },
            { name: '白色盖子', spec: '80口径', code: '820173', safeStock: 1 },
            { name: '白色盖子', spec: '75口径', code: '925667', safeStock: 1 },
            { name: '宜必思酒店盖子', spec: '75口径', code: '816561', safeStock: 1 },
            { name: '宜必思尚品酒店盖子', spec: '75口径', code: '871240', safeStock: 1 },
            { name: '星程酒店（独立包装）', spec: '9盎司', code: '3000863', safeStock: 1 },
            { name: '你好酒店2.0（独立包装）', spec: '9盎司', code: '3000864', safeStock: 1 },
            { name: '汉庭酒店（独立包装）', spec: '9盎司', code: '3000865', safeStock: 1 },
            { name: '宜必思酒店（独立包装）', spec: '9盎司', code: '3000866', safeStock: 1 },
            { name: '宜必思尚品（独立包装）', spec: '9盎司', code: '3000867', safeStock: 1 }
        ];
        const INIT_STOCK = { '849479': 452, '816545': 97, '1123064': 38, '1123065': 14, '1170947': 34, '1170948': 276,
            '955902': 49, '955901': 62, '816541': 42, '842111': 47, '842112': 29, '850078': 28, '816644': 33,
            '816645': 32, '816534': 20, '816546': 20, '1052152': 38, '816550': 65, '816552': 69, '870858': 24,
            '816547': 40, '816544': 58, '847833': 23, '816535': 27, '816525': 34, '822324': 105, '1088277': 13,
            '822325': 30, '820178': 50, '816524': 9, '816560': 0, '956922': 31, '820173': 0, '925667': 0,
            '816561': 0, '871240': 0, '3000863': 0, '3000864': 0, '3000865': 0, '3000866': 0, '3000867': 0
        };

        function createDaily() { return new Array(31).fill(0) }

        function loadData() {
            let raw = localStorage.getItem(STORAGE_KEY);
            if (raw) {
                try { data = JSON.parse(raw);
                    if (!data.products || data.products.length === 0) throw new Error('空数据'); } catch (e) {
                    data.products = PRODUCT_BASE.map((p, i) => ({ ...p, id: i + 1, initStock: INIT_STOCK[p.code] || 0,
                        dailyIn: createDaily(), dailyOut: createDaily(), shipDays: 0 }));
                    data.currentYear = 2026;
                    data.currentMonth = 6;
                    data.history = {};
                    saveData()
                }
            } else {
                data.products = PRODUCT_BASE.map((p, i) => ({ ...p, id: i + 1, initStock: INIT_STOCK[p.code] || 0,
                    dailyIn: createDaily(), dailyOut: createDaily(), shipDays: 0 }));
                data.currentYear = 2026;
                data.currentMonth = 6;
                data.history = {};
                saveData()
            }
        }

        function saveData() { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)) } catch (e) { console.warn(
                    '保存失败:', e) } }

        function calcStats(p) {
            let ti = 0,
                to = 0,
                sd = 0;
            for (let i = 0; i < 31; i++) { ti += p.dailyIn[i] || 0;
                to += p.dailyOut[i] || 0; if ((p.dailyOut[i] || 0) > 0) sd++ }
            let rs = p.initStock + ti - to,
                du = sd > 0 ? to / sd : 0,
                ad = du > 0 ? Math.floor(rs / du) : 999;
            return { totalIn: ti, totalOut: to, shipDays: sd, realStock: rs, dailyUse: du, availDays: ad }
        }

        function detectYoYAnomalies() {
            let alerts = [];
            let historyKeys = Object.keys(data.history).sort();
            if (historyKeys.length === 0) { return { alerts: [], hasData: false } }
            let lastMonthKey = historyKeys[historyKeys.length - 1];
            let lastMonthData = data.history[lastMonthKey];
            if (!lastMonthData || !lastMonthData.products) { return { alerts: [], hasData: false } }
            let lastMonthMap = {};
            lastMonthData.products.forEach(p => { lastMonthMap[p.code] = p });
            data.products.forEach(p => {
                let s = calcStats(p);
                let lastP = lastMonthMap[p.code];
                if (!lastP) return;
                let lastS = calcStats(lastP);
                if (lastS.totalOut === 0) return;
                let change = ((s.totalOut - lastS.totalOut) / lastS.totalOut) * 100;
                if (change > 50) { alerts.push({ product: p, change: change, type: 'up', current: s.totalOut,
                        last: lastS.totalOut }) } else if (change < -50) { alerts.push({ product: p,
                        change: change, type: 'down', current: s.totalOut, last: lastS.totalOut }) }
            });
            yoyAlerts = alerts;
            return { alerts, hasData: true }
        }

        function renderYoYAlerts() {
            let container = document.getElementById('yoyAlertContainer');
            let result = detectYoYAnomalies();
            if (!result.hasData) { container.innerHTML =
                    `<div style="font-size:.78em;color:var(--text2);padding:4px 0">📊 月度同比：暂无上月数据，请先完成一次月度结转。</div>`; return }
            if (result.alerts.length === 0) { container.innerHTML =
                    `<div style="font-size:.78em;color:var(--green);padding:4px 0">✅ 月度同比：各产品出库量较上月无显著异常（±50%以内）</div>`; return }
            let up = result.alerts.filter(a => a.type === 'up');
            let down = result.alerts.filter(a => a.type === 'down');
            let html =
                `<div class="alert-banner gold"><span class="label">📊 月度同比异常</span><span class="detail">${up.length}个产品增长异常，${down.length}个产品下降异常</span><button class="btn btn-sm btn-primary" onclick="showYoYDetail()">查看详情</button></div>`;
            container.innerHTML = html;
        }

        function showYoYDetail() {
            let content = document.getElementById('yoyDetailContent');
            if (yoyAlerts.length === 0) { content.innerHTML = '<p style="color:var(--text2)">暂无异常记录</p>'; return }
            let html =
                `<div class="preview-table"><table><thead><tr><th>产品</th><th>当月出库</th><th>上月出库</th><th>变化</th><th>状态</th></tr></thead><tbody>`;
            yoyAlerts.forEach(a => {
                let status = a.type === 'up' ? '<span class="tag tag-danger">↑ 增长异常</span>' :
                    '<span class="tag tag-warn">↓ 下降异常</span>';
                html +=
                    `<tr><td>${a.product.name}</td><td>${a.current}</td><td>${a.last}</td><td style="color:${a.type==='up'?'#ff6b6b':'#f0d878'};font-weight:700">${a.change > 0 ? '+':''}${a.change.toFixed(0)}%</td><td>${status}</td></tr>`;
            });
            html += `</tbody></table></div><p style="font-size:.7em;color:var(--text2);margin-top:8px">阈值：变化超过 ±50% 视为异常。</p>`;
            content.innerHTML = html;
            document.getElementById('yoyDetailModal').style.display = 'flex';
        }

        function handleSearch(value) {
            let results = document.getElementById('searchResults');
            let clearBtn = document.getElementById('searchClear');
            let trimmed = value.trim();
            if (!trimmed) { results.style.display = 'none';
                clearBtn.style.display = 'none'; return }
            clearBtn.style.display = 'inline';
            let matches = data.products.filter(p => p.name.includes(trimmed) || p.spec.includes(trimmed) || p.code
                .includes(trimmed));
            if (matches.length === 0) { results.innerHTML =
                    `<div class="item" style="color:var(--text2);cursor:default">无匹配产品</div>`;
                results.style.display = 'block'; return }
            let html = '';
            matches.slice(0, 10).forEach(p => {
                let s = calcStats(p);
                html +=
                    `<div class="item" onclick="jumpToProduct(${p.id})">${p.name} <span class="code">${p.code}</span> <span style="color:var(--text2);font-size:.7em">库存 ${s.realStock}</span></div>`;
            });
            if (matches.length > 10) { html +=
                    `<div class="item" style="color:var(--text2);cursor:default">... 共 ${matches.length} 个结果</div>` }
            results.innerHTML = html;
            results.style.display = 'block';
        }

        function jumpToProduct(pid) {
            let p = data.products.find(p => p.id === pid);
            if (!p) return;
            switchPage('stockpage');
            setTimeout(() => {
                let rows = document.querySelectorAll('#stockTbody tr');
                rows.forEach(row => {
                    row.classList.remove('highlight');
                    let nameCell = row.querySelector('td:first-child');
                    if (nameCell && nameCell.textContent.trim() === p.name) {
                        row.classList.add('highlight');
                        row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }
                });
                if (stockView === 'cards') {
                    let cards = document.querySelectorAll('#stockCardsWall .stock-card');
                    cards.forEach(card => {
                        card.style.borderColor = 'var(--border)';
                        let nameEl = card.querySelector('.sc-name');
                        if (nameEl && nameEl.textContent.trim() === p.name) {
                            card.style.borderColor = 'var(--gold)';
                            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        }
                    });
                }
            }, 200);
            clearSearch();
        }

        function clearSearch() {
            document.getElementById('searchInput').value = '';
            document.getElementById('searchResults').style.display = 'none';
            document.getElementById('searchClear').style.display = 'none';
        }

        function renderPurchaseSuggestions() {
            let container = document.getElementById('purchaseContainer');
            let suggestions = [];
            data.products.forEach(p => {
                let s = calcStats(p);
                let threshold = p.safeStock * 1.5;
                if (s.realStock < threshold && s.dailyUse > 0) {
                    let urgent = s.realStock < p.safeStock;
                    let suggestQty = Math.max(Math.ceil(p.safeStock * 2 - s.realStock), Math.ceil(s.dailyUse *
                    3));
                    let daysLeft = s.dailyUse > 0 ? Math.floor(s.realStock / s.dailyUse) : 999;
                    suggestions.push({ product: p, stats: s, urgent, suggestQty, daysLeft, threshold });
                }
            });
            suggestions.sort((a, b) => a.stats.realStock / (a.product.safeStock || 1) - b.stats.realStock / (b
                .product.safeStock || 1));
            if (suggestions.length === 0) { container.innerHTML =
                    `<div style="font-size:.8em;color:var(--green);padding:4px 0">✅ 所有产品库存充足，暂无采购建议。</div>`; return }
            let html =
                `<div style="margin-bottom:12px;font-size:.8em;color:var(--gold);font-weight:600">🛒 采购建议 · 共 ${suggestions.length} 个产品需要关注</div><div class="purchase-grid">`;
            suggestions.forEach(item => {
                let st = item.urgent ? 'pi-urgent' : 'pi-normal';
                let statusLabel = item.urgent ? '⚠️ 紧急' : '📌 关注';
                html += `<div class="purchase-item">
                    <div class="pi-name">${item.product.name} <span class="${st}" style="font-size:.7em;margin-left:6px">${statusLabel}</span></div>
                    <div class="pi-detail"><span>实时库存</span><span class="val">${item.stats.realStock}</span></div>
                    <div class="pi-detail"><span>安全库存</span><span class="val">${item.product.safeStock}</span></div>
                    <div class="pi-detail"><span>日均消耗</span><span class="val">${item.stats.dailyUse.toFixed(1)}</span></div>
                    <div class="pi-detail"><span>建议补货</span><span class="val" style="color:var(--gold2);font-weight:800">${item.suggestQty}</span></div>
                    <div class="pi-detail"><span>可支撑</span><span class="val ${item.daysLeft < 3 ? 'pi-urgent' : ''}">${item.daysLeft < 99 ? item.daysLeft+'天' : '充足'}</span></div>
                </div>`;
            });
            html += `</div>`;
            container.innerHTML = html;
        }

        function detectAnomalies() {
            let alerts = [];
            let today = new Date().getDate();
            data.products.forEach(p => {
                let s = calcStats(p);
                if (s.dailyUse <= 0) return;
                let threshold = s.dailyUse * 2.5;
                for (let i = 0; i < 31; i++) {
                    let val = p.dailyOut[i] || 0;
                    if (val > threshold && val > 5) {
                        let dayDiff = today - (i + 1);
                        if (dayDiff >= 0 && dayDiff <= 7) {
                            alerts.push({ product: p, day: i + 1, value: val, dailyUse: s.dailyUse, ratio: (val /
                                    s.dailyUse).toFixed(1) });
                        }
                    }
                }
            });
            alerts.sort((a, b) => a.day - b.day);
            detectedAlerts = alerts;
            return alerts;
        }

        function renderAlerts() {
            let container = document.getElementById('alertBannerContainer');
            let alerts = detectAnomalies();
            if (alerts.length === 0) { container.innerHTML = ''; return }
            let uniqueDays = new Set(alerts.map(a => a.day));
            let html =
                `<div class="alert-banner"><span class="label">⚠️ 异常波动报警</span><span class="detail">本月有 <strong>${uniqueDays.size}</strong> 个异常出库日（>日均2.5倍），涉及 <strong>${alerts.length}</strong> 条记录</span><button class="btn btn-sm btn-red" onclick="showAlertDetail()">查看详情</button></div>`;
            container.innerHTML = html;
        }

        function showAlertDetail() {
            let content = document.getElementById('alertDetailContent');
            if (detectedAlerts.length === 0) { content.innerHTML = '<p style="color:var(--text2)">暂无异常记录</p>'; return }
            let html =
                `<div class="preview-table"><table><thead><tr><th>日期</th><th>产品</th><th>当日出库</th><th>日均</th><th>倍数</th></tr></thead><tbody>`;
            detectedAlerts.forEach(a => {
                html +=
                    `<tr><td>${a.day}日</td><td>${a.product.name}</td><td style="color:#ff6b6b;font-weight:700">${a.value}</td><td>${a.dailyUse.toFixed(1)}</td><td>${a.ratio}x</td></tr>`;
            });
            html += `</tbody></table></div><p style="font-size:.7em;color:var(--text2);margin-top:8px">仅显示最近7天的异常记录，已自动过滤低量（<5）数据。</p>`;
            content.innerHTML = html;
            document.getElementById('alertDetailModal').style.display = 'flex';
        }

        function generateMonthlySummary() {
            let products = data.products;
            let topIn = [...products].sort((a, b) => calcStats(b).totalIn - calcStats(a).totalIn).slice(0, 3);
            let topOut = [...products].sort((a, b) => calcStats(b).totalOut - calcStats(a).totalOut).slice(0, 3);
            let topSpeed = [...products].filter(p => calcStats(p).dailyUse > 0).sort((a, b) => calcStats(b).dailyUse -
                calcStats(a).dailyUse).slice(0, 3);
            let topDrop = [...products].filter(p => calcStats(p).dailyUse > 0).sort((a, b) => {
                let sA = calcStats(a),
                    sB = calcStats(b);
                return (sA.totalIn - sA.totalOut) - (sB.totalIn - sB.totalOut);
            }).slice(0, 3);
            let lines = [];
            lines.push(`📊 ${data.currentYear}年${data.currentMonth}月 复盘摘要`);
            lines.push('');
            lines.push(`📥 入库最多：${topIn.map(p => `${p.name}(${calcStats(p).totalIn})`).join('、')}`);
            lines.push(`📤 出库最多：${topOut.map(p => `${p.name}(${calcStats(p).totalOut})`).join('、')}`);
            lines.push(`⚡ 消耗最快：${topSpeed.map(p => `${p.name}(${calcStats(p).dailyUse.toFixed(1)}/天)`).join('、')}`);
            let dropText = topDrop.map(p => { let s = calcStats(p); let net = s.totalIn - s.totalOut; return `${p.name}(${net > 0 ? '+' : ''}${net})`; })
                .join('、');
            lines.push(`📉 库存变化：${dropText || '无显著变化'}`);
            let totalIn = products.reduce((sum, p) => sum + calcStats(p).totalIn, 0);
            let totalOut = products.reduce((sum, p) => sum + calcStats(p).totalOut, 0);
            let netChange = totalIn - totalOut;
            let totalStock = products.reduce((sum, p) => sum + calcStats(p).realStock, 0);
            lines.push(
                `📦 总览：入库${totalIn}，出库${totalOut}，净${netChange > 0 ? '+' : ''}${netChange}，期末库存${totalStock}`);
            lines.push('');
            lines.push(`—— 婉玲自动生成 · ${new Date().toLocaleString()}`);
            return lines.join('\n');
        }

        function checkMonthRollover() {
            let now = new Date();
            let currentYear = now.getFullYear();
            let currentMonth = now.getMonth() + 1;
            let needRollover = false;
            let msg = '';
            if (data.currentYear < currentYear) {
                needRollover = true;
                msg = `系统记录为 ${data.currentYear}年${data.currentMonth}月，当前已进入 ${currentYear}年${currentMonth}月。`;
            } else if (data.currentYear === currentYear && data.currentMonth < currentMonth) {
                needRollover = true;
                msg = `系统记录为 ${data.currentYear}年${data.currentMonth}月，当前已进入 ${currentYear}年${currentMonth}月。`;
            } else if (data.currentYear === currentYear && data.currentMonth === currentMonth) {
                msg = `系统时间与当前月份一致 (${data.currentYear}年${data.currentMonth}月)，无需结转。`;
            }
            if (needRollover) {
                document.getElementById('carryoverMsg').textContent =
                    `${msg} 是否自动结转 ${data.currentYear}年${data.currentMonth}月的数据并切换至 ${currentYear}年${currentMonth}月？`;
                let summary = generateMonthlySummary();
                document.getElementById('carryoverSummary').textContent = summary;
                document.getElementById('carryoverSummary').style.display = 'block';
                document.getElementById('carryoverModal').style.display = 'flex';
                window._rolloverTarget = { year: currentYear, month: currentMonth };
            } else { alert(msg || '当前月份无需结转。') }
        }

        function monthlyCarryover() {
            document.getElementById('carryoverMsg').textContent =
                `即将把${data.currentYear}年${data.currentMonth}月的实时库存结转为下月盘点库存，本月数据自动存档至历史记录。`;
            let summary = generateMonthlySummary();
            document.getElementById('carryoverSummary').textContent = summary;
            document.getElementById('carryoverSummary').style.display = 'block';
            document.getElementById('carryoverModal').style.display = 'flex';
            window._rolloverTarget = null;
        }

        function confirmCarryover() {
            let targetYear = window._rolloverTarget ? window._rolloverTarget.year : data.currentYear;
            let targetMonth = window._rolloverTarget ? window._rolloverTarget.month : data.currentMonth + 1;
            if (targetMonth > 12) { targetMonth = 1;
                targetYear++ }
            let key = `${data.currentYear}年${data.currentMonth}月`;
            let snapshot = JSON.parse(JSON.stringify(data.products));
            let summary = generateMonthlySummary();
            data.history[key] = { products: snapshot, summary: summary };
            data.products.forEach(p => { let s = calcStats(p);
                p.initStock = s.realStock;
                p.dailyIn = createDaily();
                p.dailyOut = createDaily();
                p.shipDays = 0 });
            data.currentYear = targetYear;
            data.currentMonth = targetMonth;
            saveData();
            renderAll();
            document.getElementById('carryoverModal').style.display = 'none';
            document.getElementById('carryoverSummary').style.display = 'none';
            window._rolloverTarget = null;
            alert(`结转完成！${key}数据已存档，已切换至 ${targetYear}年${targetMonth}月。\n\n${summary}`);
        }

        function renderHistoryList() {
            let container = document.getElementById('historyList');
            let keys = Object.keys(data.history).sort().reverse();
            if (keys.length === 0) { container.innerHTML =
                    '<span style="color:#8899aa;font-size:.8em">暂无历史记录，月度结转后自动保存</span>'; return }
            container.innerHTML = keys.map(k =>
                `<div class="history-chip" onclick="viewHistory('${k}')">📅 ${k}</div>`).join('');
            if (keys.length > 0) { viewHistory(keys[0]); }
        }

        function viewHistory(key) {
            let h = data.history[key];
            if (!h) return;
            document.querySelectorAll('.history-chip').forEach(c => c.classList.remove('active'));
            event.target.classList.add('active');
            let summaryContainer = document.getElementById('historySummaryContainer');
            if (h.summary) { summaryContainer.innerHTML =
                    `<div class="summary-text"><strong>📋 ${key} 复盘摘要</strong>\n${h.summary}</div>` } else { summaryContainer
                    .innerHTML = '' }
            let tb = document.getElementById('historyTbody');
            tb.innerHTML = h.products.map(p => { let s = calcStats(p);
                return `<tr><td style="text-align:left;font-weight:600">${p.name}</td><td>${p.spec}</td><td>${p.code}</td><td>${p.initStock}</td><td>${p.safeStock}</td>` +
                    p.dailyOut.map(v => `<td>${v||''}</td>`).join('') +
                    `<td>${s.totalIn}</td><td>${s.totalOut}</td><td>${s.shipDays}天</td><td>${s.dailyUse.toFixed(1)}</td><td>${s.realStock}</td><td>${s.dailyUse>0?s.availDays+'天':'-'}</td></tr>`
            }).join('');
        }

        function renderAll() {
            document.getElementById('monthBadge').textContent = `📅 ${data.currentYear}年${data.currentMonth}月`;
            renderStats();
            renderInTable();
            renderOutTable();
            renderStockPage();
            renderCharts();
            renderHistoryList();
            renderProductSelect();
            renderPurchaseSuggestions();
            renderAlerts();
            renderYoYAlerts();
        }

        function renderStats() {
            let ts = 0,
                lc = 0,
                oc = 0,
                totalDailyUse = 0;
            data.products.forEach(p => { let s = calcStats(p);
                ts += s.realStock;
                totalDailyUse += s.dailyUse; if (s.realStock <= p.safeStock && s.realStock > 0) lc++; if (s.realStock <=
                    0) oc++ });
            let health = totalDailyUse > 0 ? Math.floor(ts / totalDailyUse) : 999;
            let healthColor = health < 3 ? 'red' : health < 7 ? 'yellow' : 'green';
            let healthLabel = health < 3 ? '⚠️ 库存紧张' : health < 7 ? '⚡ 需关注' : '✅ 健康';
            document.getElementById('stats').innerHTML =
                `<div class="stat-card"><div class="icon">📋</div><div class="num">${data.products.length}</div><div class="label">产品总数</div></div><div class="stat-card"><div class="icon">📦</div><div class="num">${ts.toLocaleString()}</div><div class="label">库存总量</div></div><div class="stat-card"><div class="icon">⚠️</div><div class="num warn">${lc}</div><div class="label">低库存预警</div></div><div class="stat-card"><div class="icon">🚫</div><div class="num warn">${oc}</div><div class="label">已断货</div></div><div class="stat-card"><div class="icon">❤️</div><div class="num ${healthColor}">${health < 99 ? health+'天' : '∞'}</div><div class="label">库存健康度 ${healthLabel}</div><div class="sub">总库存 / 日均消耗</div></div>`;
        }

        function getToday() { let d = new Date(); return d.getDate() }

        function addTodayIn(pid) {
            let p = data.products.find(p => p.id === pid);
            if (!p) return;
            let today = getToday() - 1;
            if (today < 0 || today > 30) return alert('今日日期超出范围');
            p.dailyIn[today] = (p.dailyIn[today] || 0) + 1;
            saveData();
            renderAll();
        }

        function addTodayOut(pid) {
            let p = data.products.find(p => p.id === pid);
            if (!p) return;
            let today = getToday() - 1;
            if (today < 0 || today > 30) return alert('今日日期超出范围');
            p.dailyOut[today] = (p.dailyOut[today] || 0) + 1;
            saveData();
            renderAll();
        }

        function quickBatch() {
            let input = document.getElementById('quickInput').value.trim();
            let result = document.getElementById('quickResult');
            if (!input) { result.textContent = '⚠️ 请输入指令'; return }
            let parts = input.split(/[\s,，、\n]+/).filter(s => s);
            let matched = 0;
            let errors = [];
            let idx = 0;
            while (idx < parts.length) {
                let namePart = '';
                let op = null;
                let qty = null;
                let found = false;
                for (let end = idx + 1; end <= Math.min(idx + 6, parts.length); end++) {
                    let testName = parts.slice(idx, end).join(' ');
                    let remaining = parts.slice(end);
                    for (let j = 0; j < remaining.length; j++) {
                        let opCandidate = remaining[j];
                        let opType = null;
                        if (opCandidate === '入库' || opCandidate === '入') opType = 'in';
                        else if (opCandidate === '出库' || opCandidate === '出') opType = 'out';
                        else if (opCandidate === '进' || opCandidate === '销') continue;
                        if (opType) {
                            let qtyVal = parseInt(remaining[j + 1]);
                            if (!isNaN(qtyVal) && qtyVal > 0) {
                                let p = data.products.find(p => p.name.includes(testName) || p.spec.includes(
                                    testName) || p.code === testName);
                                if (p) {
                                    namePart = testName;
                                    op = opType;
                                    qty = qtyVal;
                                    idx = end;
                                    idx += 2;
                                    found = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (found) break;
                }
                if (!found) {
                    let fullMatch = false;
                    for (let p of data.products) {
                        let pattern = p.name + ' 入库 ';
                        if (input.includes(pattern)) {
                            let parts2 = input.split(pattern);
                            if (parts2.length > 1) {
                                let num = parseInt(parts2[1]);
                                if (!isNaN(num) && num > 0) {
                                    let qtyVal = num;
                                    let target = p;
                                    let today = getToday() - 1;
                                    if (today < 0 || today > 30) { errors.push('日期超出'); break }
                                    target.dailyIn[today] = (target.dailyIn[today] || 0) + qtyVal;
                                    matched++;
                                    input = input.replace(pattern + num, '');
                                    fullMatch = true;
                                    break;
                                }
                            }
                        }
                        let pattern2 = p.name + ' 出库 ';
                        if (input.includes(pattern2)) {
                            let parts2 = input.split(pattern2);
                            if (parts2.length > 1) {
                                let num = parseInt(parts2[1]);
                                if (!isNaN(num) && num > 0) {
                                    let qtyVal = num;
                                    let target = p;
                                    let today = getToday() - 1;
                                    if (today < 0 || today > 30) { errors.push('日期超出'); break }
                                    target.dailyOut[today] = (target.dailyOut[today] || 0) + qtyVal;
                                    matched++;
                                    input = input.replace(pattern2 + num, '');
                                    fullMatch = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (!fullMatch) {
                        errors.push(`无法识别: ${parts.slice(idx, Math.min(idx+5, parts.length)).join(' ')}`);
                        idx += 1;
                    }
                } else {
                    let target = data.products.find(p => p.name.includes(namePart) || p.spec.includes(namePart) || p
                        .code === namePart);
                    if (target && op && qty) {
                        let today = getToday() - 1;
                        if (today < 0 || today > 30) { errors.push('日期超出'); continue }
                        if (op === 'in') { target.dailyIn[today] = (target.dailyIn[today] || 0) + qty; } else { target
                                .dailyOut[today] = (target.dailyOut[today] || 0) + qty; }
                        matched++;
                    } else { errors.push(`匹配失败: ${namePart}`); }
                }
                if (idx >= parts.length) break;
            }
            saveData();
            renderAll();
            if (matched > 0) {
                result.textContent = `✅ 成功录入 ${matched} 条指令${errors.length > 0 ? '，'+errors.length+'条失败' : ''}`;
            } else { result.textContent =
                    `❌ 未能识别任何指令。请尝试：产品名 入库/出库 数量，例如：全季6.5盎司 入库 50` }
            if (errors.length > 0) { console.warn('批量录入错误:', errors) }
        }

        function renderInTable() {
            let tb = document.getElementById('inTbody');
            tb.innerHTML = data.products.map(p => {
                let s = calcStats(p);
                let h =
                    `<tr><td style="text-align:left;font-weight:600">${p.name}</td><td>${p.spec}</td><td style="font-family:monospace;color:var(--gold)">${p.code}</td><td>${p.initStock}</td><td>${p.safeStock}</td>`;
                for (let i = 0; i < 31; i++) {
                    h +=
                        `<td><input class="cell-input" type="number" value="${p.dailyIn[i]||''}" onchange="updateCell(${p.id},'in',${i},this.value)"></td>`;
                }
                h += `<td class="cell-sum in-sum">${s.totalIn}</td>`;
                h += `<td><button class="btn-today" onclick="addTodayIn(${p.id})">+今日</button></td>`;
                h += `</tr>`;
                return h;
            }).join('')
        }

        function renderOutTable() {
            let tb = document.getElementById('outTbody');
            tb.innerHTML = data.products.map(p => {
                let s = calcStats(p);
                let st = s.realStock <= 0 ? '<span class="tag tag-danger">断货</span>' : s.realStock <= p
                    .safeStock ? '<span class="tag tag-warn">不足</span>' :
                    '<span class="tag tag-ok">正常</span>';
                let h =
                    `<tr><td style="text-align:left;font-weight:600">${p.name}</td><td>${p.spec}</td><td style="font-family:monospace;color:var(--gold)">${p.code}</td>`;
                for (let i = 0; i < 31; i++) {
                    h +=
                        `<td><input class="cell-input" type="number" value="${p.dailyOut[i]||''}" onchange="updateCell(${p.id},'out',${i},this.value)"></td>`;
                }
                h +=
                    `<td class="cell-sum out-sum">${s.totalOut}</td><td>${s.shipDays}天</td><td>${s.dailyUse.toFixed(1)}</td><td style="color:${s.realStock<=p.safeStock?'var(--red)':'var(--green)'};font-weight:700">${s.realStock}</td><td>${s.dailyUse>0?s.availDays+'天':'-'}</td><td>${st}</td>`;
                h += `<td><button class="btn-today btn-today-out" onclick="addTodayOut(${p.id})">+今日</button></td>`;
                h += `</tr>`;
                return h;
            }).join('')
        }

        function renderStockPage() { renderStockCards();
            renderStockTable() }

        function renderStockCards() {
            let wall = document.getElementById('stockCardsWall');
            let sorted = [...data.products].sort((a, b) => calcStats(a).realStock - calcStats(b).realStock);
            wall.innerHTML = sorted.map(p => { let s = calcStats(p);
                let sc = s.realStock <= 0 ? 'status-danger' : s.realStock <= p.safeStock ? 'status-warn' :
                'status-ok';
                let st = s.realStock <= 0 ? '断货' : s.realStock <= p.safeStock ? '不足' : '正常';
                let sc2 = s.realStock <= 0 ? 'color:#ff6b6b' : s.realStock <= p.safeStock ? 'color:#f0d878' :
                    'color:#4a9e6e';
                return `<div class="stock-card"><span class="sc-status ${sc}">${st}</span><div class="sc-name">${p.name}</div><div class="sc-spec">${p.code} · ${p.spec}</div><div class="sc-stock" style="${sc2}">${s.realStock}</div><div class="sc-info"><div>可发<span>${s.dailyUse>0?s.availDays+'天':'-'}</span></div><div>日均<span>${s.dailyUse.toFixed(1)}</span></div></div><div class="sc-safe">🔒 安全库存 <input type="number" value="${p.safeStock}" onchange="updateSafeStock(${p.id},this.value)"></div></div>`
            }).join('')
        }

        function renderStockTable() {
            let tb = document.getElementById('stockTbody');
            let sorted = [...data.products].sort((a, b) => calcStats(a).realStock - calcStats(b).realStock);
            tb.innerHTML = sorted.map(p => { let s = calcStats(p);
                let sc = s.realStock <= 0 ? 'color:#ff6b6b' : s.realStock <= p.safeStock ? 'color:#f0d878' :
                    'color:#4a9e6e';
                let st = s.realStock <= 0 ? '<span class="tag tag-danger">断货</span>' : s.realStock <= p
                    .safeStock ? '<span class="tag tag-warn">不足</span>' :
                    '<span class="tag tag-ok">正常</span>';
                return `<tr><td style="text-align:left;font-weight:600">${p.name}</td><td>${p.spec}</td><td style="font-family:monospace;color:var(--gold)">${p.code}</td><td>${p.initStock}</td><td style="${sc};font-weight:700">${s.realStock}</td><td><input class="cell-input-safe" type="number" value="${p.safeStock}" onchange="updateSafeStock(${p.id},this.value)"></td><td>${s.dailyUse>0?s.availDays+'天':'-'}</td><td>${s.dailyUse.toFixed(1)}</td><td>${st}</td></tr>`
            }).join('')
        }

        function updateSafeStock(pid, value) { let p = data.products.find(p => p.id === pid); if (!p) return;
            p.safeStock = parseInt(value) || 0;
            saveData();
            renderStockPage();
            renderStats();
            renderPurchaseSuggestions(); }

        function switchStockView(view) {
            stockView = view;
            document.getElementById('stockCardsView').style.display = view === 'cards' ? 'block' : 'none';
            document.getElementById('stockTableView').style.display = view === 'table' ? 'block' : 'none';
            document.getElementById('tabStockCards').classList.toggle('active', view === 'cards');
            document.getElementById('tabStockTable').classList.toggle('active', view === 'table')
        }

        function initStockView() {
            stockView = 'cards';
            document.getElementById('stockCardsView').style.display = 'block';
            document.getElementById('stockTableView').style.display = 'none';
            document.getElementById('tabStockCards').classList.add('active');
            document.getElementById('tabStockTable').classList.remove('active')
        }

        function updateCell(pid, type, index, value) { let p = data.products.find(p => p.id === pid); if (!p) return; if (
                type === 'in') { p.dailyIn[index] = parseInt(value) || 0 } else { p.dailyOut[index] = parseInt(
                    value) || 0 }
            saveData();
            renderAll() }

        function renderProductSelect() {
            let sel = document.getElementById('singleProductSelect');
            if (!sel) return;
            sel.innerHTML = data.products.map(p => `<option value="${p.id}">${p.name} (${p.spec})</option>`).join('');
            updateSingleChart()
        }

        function updateSingleChart() {
            let sel = document.getElementById('singleProductSelect');
            if (!sel) return;
            let pid = parseInt(sel.value);
            let p = data.products.find(p => p.id === pid);
            if (!p) return;
            let labels = [];
            for (let i = 1; i <= 31; i++) labels.push(i + '日');
            let sin = [],
                sout = [];
            for (let i = 0; i < 31; i++) { sin.push(p.dailyIn[i] || 0);
                sout.push(p.dailyOut[i] || 0) }
            if (singleChart) singleChart.destroy();
            singleChart = new Chart(document.getElementById('singleChart'), { type: 'line', data: { labels,
                    datasets: [{ label: p.name + ' 入库', data: sin, borderColor: '#3fb950',
                            backgroundColor: 'rgba(63,185,80,0.06)', fill: true, tension: 0.35,
                            pointRadius: 0, borderWidth: 2 }, { label: p.name + ' 出库', data: sout,
                            borderColor: '#e94560', backgroundColor: 'rgba(233,69,96,0.06)',
                            fill: true, tension: 0.35, pointRadius: 0, borderWidth: 2 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8899aa',
                            font: { size: 10 }, usePointStyle: true } } },
                    scales: { x: { ticks: { color: '#556', maxTicksLimit: 15 }, grid: { color: 'rgba(255,255,255,0.03)' } },
                        y: { ticks: { color: '#556' }, grid: { color: 'rgba(255,255,255,0.03)' } } } } });
            if (singleBarChart) singleBarChart.destroy();
            singleBarChart = new Chart(document.getElementById('singleBarChart'), { type: 'bar', data: { labels,
                    datasets: [{ label: p.name + ' 入库', data: sin, backgroundColor: 'rgba(74,158,110,0.7)',
                            borderRadius: 4, borderWidth: 0 }, { label: p.name + ' 出库', data: sout,
                            backgroundColor: 'rgba(212,85,74,0.7)', borderRadius: 4, borderWidth: 0 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8899aa',
                            font: { size: 10 }, usePointStyle: true } } },
                    scales: { x: { ticks: { color: '#556', maxTicksLimit: 15 }, grid: { color: 'rgba(255,255,255,0.03)' } },
                        y: { ticks: { color: '#556' }, grid: { color: 'rgba(255,255,255,0.03)' } } } } })
        }

        function renderCharts() {
            let labels = [];
            for (let i = 1; i <= 31; i++) labels.push(i + '日');
            let diT = new Array(31).fill(0),
                doT = new Array(31).fill(0);
            data.products.forEach(p => { for (let i = 0; i < 31; i++) { diT[i] += p.dailyIn[i] || 0;
                    doT[i] += p.dailyOut[i] || 0 } });
            if (dailyChart) dailyChart.destroy();
            dailyChart = new Chart(document.getElementById('dailyChart'), { type: 'line', data: { labels,
                    datasets: [{ label: '入库', data: diT, borderColor: '#3fb950',
                            backgroundColor: 'rgba(63,185,80,0.06)', fill: true, tension: 0.35,
                            pointRadius: 0, borderWidth: 2 }, { label: '出库', data: doT,
                            borderColor: '#e94560', backgroundColor: 'rgba(233,69,96,0.06)',
                            fill: true, tension: 0.35, pointRadius: 0, borderWidth: 2 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8899aa',
                            font: { size: 10 }, usePointStyle: true } } },
                    scales: { x: { ticks: { color: '#556', maxTicksLimit: 15 }, grid: { color: 'rgba(255,255,255,0.03)' } },
                        y: { ticks: { color: '#556' }, grid: { color: 'rgba(255,255,255,0.03)' } } } } });
            let allIn = data.products.map(p => calcStats(p).totalIn),
                allOut = data.products.map(p => calcStats(p).totalOut),
                allLabels = data.products.map(p => p.name);
            if (barAllChart) barAllChart.destroy();
            barAllChart = new Chart(document.getElementById('barAllChart'), { type: 'bar', data: { labels: allLabels,
                    datasets: [{ label: '月入库', data: allIn, backgroundColor: 'rgba(74,158,110,0.7)',
                            borderRadius: 6, borderWidth: 0, barPercentage: 0.55, categoryPercentage: 0.7 },
                        { label: '月出库', data: allOut, backgroundColor: 'rgba(212,85,74,0.7)',
                            borderRadius: 6, borderWidth: 0, barPercentage: 0.55,
                        categoryPercentage: 0.7 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8899aa',
                            font: { size: 10 }, usePointStyle: true, padding: 15 } } },
                    scales: { x: { ticks: { color: '#8899aa', font: { size: 7 }, maxRotation: 60,
                                autoSkip: false }, grid: { display: false } }, y: { ticks: { color: '#556' },
                            grid: { color: 'rgba(255,255,255,0.04)' }, beginAtZero: true } } } });
            if (alertChart) alertChart.destroy();
            let safe = data.products.filter(p => calcStats(p).realStock > p.safeStock).length,
                low = data.products.filter(p => calcStats(p).realStock <= p.safeStock && calcStats(p)
                    .realStock > 0).length,
                out = data.products.filter(p => calcStats(p).realStock <= 0).length;
            alertChart = new Chart(document.getElementById('alertChart'), { type: 'doughnut', data: { labels: [
                        '库存充足', '低库存', '已断货'
                    ], datasets: [{ data: [safe, low, out],
                        backgroundColor: ['#4a9e6e', '#c8a84e', '#d4554a'], borderWidth: 2,
                        borderColor: '#1a1f2e' }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#8899aa',
                            font: { size: 10 }, usePointStyle: true, padding: 18 } } },
                    cutout: '65%' } })
        }

        function switchPage(page) {
            currentPage = page;
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            document.getElementById('page-' + page).classList.add('active');
            document.querySelectorAll('.menu-item').forEach(m => m.classList.remove('active'));
            document.querySelector(`[data-page="${page}"]`).classList.add('active');
            if (page === 'stockpage') { initStockView() }
            renderAll()
        }

        function exportExcel() {
            let rows = [
                ['产品名称', '规格', '编码', '盘点库存', '安全线', '月入库', '月出库', '发货天数', '日均消耗', '实时库存', '可发天数',
                    '状态'
                ]
            ];
            data.products.forEach(p => { let s = calcStats(p);
                let st = s.realStock <= 0 ? '断货' : s.realStock <= p.safeStock ? '不足' : '正常';
                rows.push([p.name, p.spec, p.code, p.initStock, p.safeStock, s.totalIn, s.totalOut, s.shipDays,
                    parseFloat(s.dailyUse.toFixed(1)), s.realStock, s.dailyUse > 0 ? s.availDays : '', st
                ]) });
            let ws = XLSX.utils.aoa_to_sheet(rows);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '进销存');
            XLSX.writeFile(wb, `进销存_${data.currentYear}年${data.currentMonth}月.xlsx`)
        }

        function exportInExcel() {
            let rows = [
                ['产品名称', '规格', '编码', '盘点库存', '安全线']
            ];
            for (let i = 1; i <= 31; i++) rows[0].push(i + '日');
            rows[0].push('月汇总');
            data.products.forEach(p => { let s = calcStats(p);
                let row = [p.name, p.spec, p.code, p.initStock, p.safeStock]; for (let i = 0; i < 31; i++) row
                    .push(p.dailyIn[i] || 0);
                row.push(s.totalIn);
                rows.push(row) });
            let ws = XLSX.utils.aoa_to_sheet(rows);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '月入库明细');
            XLSX.writeFile(wb, `月入库明细_${data.currentYear}年${data.currentMonth}月.xlsx`)
        }

        function exportOutExcel() {
            let rows = [
                ['产品名称', '规格', '编码']
            ];
            for (let i = 1; i <= 31; i++) rows[0].push(i + '日');
            rows[0].push('月汇总', '发货天数', '日均消耗', '实时库存', '可发天数', '状态');
            data.products.forEach(p => { let s = calcStats(p);
                let st = s.realStock <= 0 ? '断货' : s.realStock <= p.safeStock ? '不足' : '正常';
                let row = [p.name, p.spec, p.code]; for (let i = 0; i < 31; i++) row.push(p.dailyOut[i] || 0);
                row.push(s.totalOut, s.shipDays, parseFloat(s.dailyUse.toFixed(1)), s.realStock, s.dailyUse > 0 ?
                    s.availDays : '', st);
                rows.push(row) });
            let ws = XLSX.utils.aoa_to_sheet(rows);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '月出库明细');
            XLSX.writeFile(wb, `月出库明细_${data.currentYear}年${data.currentMonth}月.xlsx`)
        }

        function importDailyExcel(file, type) {
            let reader = new FileReader();
            reader.onload = function(ev) {
                try {
                    let wb = XLSX.read(ev.target.result, { type: 'array' });
                    let ws = wb.Sheets[wb.SheetNames[0]];
                    let rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
                    if (rows.length < 2) return alert('Excel至少需要标题行+一行数据');
                    let header = rows[0];
                    let codeCol = -1,
                        firstDayCol = -1;
                    for (let i = 0; i < header.length; i++) {
                        let h = String(header[i]).trim();
                        if ((h === '编码' || h === '产品编码') && codeCol === -1) codeCol = i;
                        if (h === '1日' && firstDayCol === -1) firstDayCol = i
                    }
                    if (codeCol === -1) return alert('缺少编码列');
                    if (firstDayCol === -1) return alert('缺少1日列');
                    let preview = [];
                    let matchedCodes = [];
                    for (let i = 1; i < rows.length; i++) {
                        let row = rows[i];
                        if (!row || !row[codeCol]) continue;
                        let code = String(row[codeCol]).trim();
                        let p = data.products.find(p => p.code === code);
                        let updateCount = 0;
                        if (p) {
                            for (let d = 0; d < 31; d++) {
                                let colIdx = firstDayCol + d;
                                if (colIdx < row.length && row[colIdx] !== undefined && row[colIdx] !== '') {
                                    let val = parseInt(row[colIdx]);
                                    if (!isNaN(val)) updateCount++;
                                }
                            }
                        }
                        preview.push({ code, name: p ? p.name : '未找到', spec: p ? p.spec : '-', updateCount,
                            exists: !!p });
                        if (p) matchedCodes.push(code);
                    }
                    let content = document.getElementById('importPreviewContent');
                    let html =
                        `<div class="preview-table"><table><thead><tr><th>编码</th><th>产品</th><th>规格</th><th>待更新数量</th><th>状态</th></tr></thead><tbody>`;
                    preview.forEach(item => {
                        let status = item.exists ? '<span class="tag tag-ok">匹配</span>' :
                            '<span class="tag tag-danger">未匹配</span>';
                        html +=
                            `<tr><td>${item.code}</td><td>${item.name}</td><td>${item.spec}</td><td>${item.updateCount}</td><td>${status}</td></tr>`;
                    });
                    html +=
                        `</tbody></table></div><p style="font-size:.78em;color:var(--text2);margin-top:8px">共 ${preview.length} 行，其中 ${matchedCodes.length} 个产品匹配成功。未匹配的产品不会被导入。</p>`;
                    content.innerHTML = html;
                    pendingImportData = { type, rows, codeCol, firstDayCol, matchedCodes };
                    document.getElementById('importPreviewModal').style.display = 'flex';
                } catch (err) { alert('Excel解析失败，请检查文件格式') }
            };
            reader.readAsArrayBuffer(file)
        }

        function handleInExcelImport(e) { if (!e.target.files[0]) return;
            importDailyExcel(e.target.files[0], 'in');
            e.target.value = '' }

        function handleOutExcelImport(e) { if (!e.target.files[0]) return;
            importDailyExcel(e.target.files[0], 'out');
            e.target.value = '' }

        function handleJsonImport(e) {
            let f = e.target.files[0];
            if (!f) return;
            let reader = new FileReader();
            reader.onload = function(ev) {
                try {
                    let imp = JSON.parse(ev.target.result);
                    if (!imp.products) return alert('格式错误，缺少products字段');
                    let preview = imp.products.map(p => ({
                        code: p.code,
                        name: p.name || '未命名',
                        spec: p.spec || '-',
                        exists: data.products.some(dp => dp.code === p.code)
                    }));
                    let content = document.getElementById('importPreviewContent');
                    let html =
                        `<div class="preview-table"><table><thead><tr><th>编码</th><th>产品</th><th>规格</th><th>状态</th></tr></thead><tbody>`;
                    preview.forEach(item => {
                        let status = item.exists ? '<span class="tag tag-ok">已存在</span>' :
                            '<span class="tag tag-warn">新产品</span>';
                        html +=
                            `<tr><td>${item.code}</td><td>${item.name}</td><td>${item.spec}</td><td>${status}</td></tr>`;
                    });
                    html +=
                        `</tbody></table></div><p style="font-size:.78em;color:var(--text2);margin-top:8px">共 ${preview.length} 个产品。${preview.filter(p=>p.exists).length} 个已存在，${preview.filter(p=>!p.exists).length} 个新产品。</p>`;
                    content.innerHTML = html;
                    pendingImportData = { type: 'json', data: imp };
                    document.getElementById('importPreviewModal').style.display = 'flex';
                } catch (err) { alert('JSON解析失败: ' + err.message) }
            };
            reader.readAsText(f);
            e.target.value = ''
        }

        function confirmImport() {
            if (!pendingImportData) return;
            if (pendingImportData.type === 'json') {
                let imp = pendingImportData.data;
                if (!confirm(`导入 ${imp.products.length} 个产品，覆盖当前数据？`)) { document.getElementById('importPreviewModal')
                        .style.display = 'none';
                    pendingImportData = null; return }
                data.products = imp.products;
                if (imp.currentYear) data.currentYear = imp.currentYear;
                if (imp.currentMonth) data.currentMonth = imp.currentMonth;
                if (imp.history) data.history = imp.history;
                saveData();
                renderAll();
                alert('导入成功');
                document.getElementById('importPreviewModal').style.display = 'none';
                pendingImportData = null;
                return;
            }
            let { type, rows, codeCol, firstDayCol, matchedCodes } = pendingImportData;
            let updatedCount = 0;
            for (let i = 1; i < rows.length; i++) {
                let row = rows[i];
                if (!row || !row[codeCol]) continue;
                let code = String(row[codeCol]).trim();
                let p = data.products.find(p => p.code === code);
                if (!p) continue;
                for (let d = 0; d < 31; d++) {
                    let colIdx = firstDayCol + d;
                    if (colIdx < row.length && row[colIdx] !== undefined && row[colIdx] !== '') {
                        let val = parseInt(row[colIdx]);
                        if (!isNaN(val)) {
                            if (type === 'in') p.dailyIn[d] = val;
                            else p.dailyOut[d] = val;
                            updatedCount++;
                        }
                    }
                }
            }
            if (updatedCount > 0) {
                saveData();
                renderAll();
                alert(`导入成功！更新了 ${updatedCount} 个日度数据。`);
            } else { alert('没有数据被更新') }
            document.getElementById('importPreviewModal').style.display = 'none';
            pendingImportData = null;
        }

        function exportData() {
            let exp = { products: data.products, currentYear: data.currentYear, currentMonth: data.currentMonth,
                history: data.history, exportTime: new Date().toISOString() };
            let blob = new Blob([JSON.stringify(exp, null, 2)], { type: 'application/json' });
            let a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `进销存备份_${data.currentYear}年${data.currentMonth}月.json`;
            a.click()
        }

        function openAddProduct(m, pid) {
            if (!m) m = 'add';
            document.getElementById('modalTitle').textContent = m === 'edit' ? '✏️ 编辑产品' : '➕ 新增产品';
            if (m === 'edit' && pid) {
                let p = data.products.find(p => p.id === pid);
                if (p) {
                    document.getElementById('mName').value = p.name;
                    document.getElementById('mSpec').value = p.spec;
                    document.getElementById('mCode').value = p.code;
                    document.getElementById('mSafeStock').value = p.safeStock;
                    document.getElementById('mInitStock').value = p.initStock;
                    document.getElementById('mEditId').value = p.id
                }
            } else {
                document.getElementById('mName').value = '';
                document.getElementById('mSpec').value = '';
                document.getElementById('mCode').value = '';
                document.getElementById('mSafeStock').value = '10';
                document.getElementById('mInitStock').value = '0';
                document.getElementById('mEditId').value = ''
            }
            document.getElementById('addProductModal').style.display = 'flex'
        }

        function closeAddProduct() { document.getElementById('addProductModal').style.display = 'none' }

        function saveProduct() {
            let n = document.getElementById('mName').value.trim(),
                s = document.getElementById('mSpec').value.trim(),
                c = document.getElementById('mCode').value.trim(),
                sf = parseInt(document.getElementById('mSafeStock').value) || 10,
                ist = parseInt(document.getElementById('mInitStock').value) || 0,
                eid = document.getElementById('mEditId').value;
            if (!n || !c) return alert('产品名称和编码必填');
            if (eid) {
                let p = data.products.find(p => p.id == eid);
                if (p) { p.name = n;
                    p.spec = s;
                    p.code = c;
                    p.safeStock = sf; if (ist != p.initStock) p.initStock = ist }
            } else data.products.push({ id: Date.now(), name: n, spec: s, code: c, safeStock: sf, initStock: ist,
                dailyIn: createDaily(), dailyOut: createDaily(), shipDays: 0 });
            closeAddProduct();
            saveData();
            renderAll()
        }

        loadData();
        renderAll();
        setTimeout(() => { checkMonthRollover(); }, 300);
    </script>
</body>
</html>
