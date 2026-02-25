#!/usr/bin/env python3
"""Build an industry-standard competitive benchmark XLSX (5 sheets) from JSON input.

Supports multilingual output based on explicit --lang or automatic language detection.
No third-party dependency required.
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from datetime import date
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

SHEET_ORDER = ["summary", "benchmark", "feature_matrix", "pricing_gtm", "sources"]

SHEET_LAYOUT: dict[str, dict[str, Any]] = {
    "summary": {
        "columns": [
            "problem_statement",
            "target_segment",
            "method",
            "scope",
            "top_findings",
            "strategic_implications",
        ],
        "widths": [34, 24, 18, 20, 48, 48],
    },
    "benchmark": {
        "columns": [
            "rank",
            "company_product",
            "category",
            "target_user",
            "core_jtbd",
            "platform",
            "geo_focus",
            "traction_score",
            "product_capability_score",
            "monetization_score",
            "user_sentiment_score",
            "execution_maturity_score",
            "evidence_confidence_score",
            "weighted_total",
            "key_strength",
            "key_weakness",
            "threat_level",
        ],
        "widths": [8, 26, 30, 20, 26, 16, 14, 16, 20, 18, 18, 20, 20, 18, 28, 28, 14],
    },
    "feature_matrix": {
        "columns": [
            "l1_capability",
            "l2_module",
            "l3_feature",
            "our_status",
            "competitor_coverage",
            "parity_gap",
            "importance",
            "priority",
        ],
        "widths": [24, 22, 34, 30, 24, 18, 18, 18],
    },
    "pricing_gtm": {
        "columns": [
            "product",
            "pricing_model",
            "entry_price",
            "top_tier_price",
            "trial_freemium",
            "packaging_unit",
            "primary_channel",
            "positioning_claim",
            "observed_conversion_frictions",
        ],
        "widths": [22, 20, 14, 14, 18, 18, 32, 36, 36],
    },
    "sources": {
        "columns": [
            "product",
            "source_type",
            "url",
            "title",
            "published_date",
            "access_date",
            "claim",
            "evidence_snippet",
            "confidence",
        ],
        "widths": [24, 28, 46, 32, 16, 14, 34, 46, 18],
    },
}

LOCALES: dict[str, dict[str, Any]] = {
    "en": {
        "sheet_names": {
            "summary": "Summary",
            "benchmark": "Benchmark",
            "feature_matrix": "Feature-Matrix",
            "pricing_gtm": "Pricing-GTM",
            "sources": "Sources",
        },
        "headers": {
            "problem_statement": "Problem Statement",
            "target_segment": "Target Segment",
            "method": "Method",
            "scope": "Scope",
            "top_findings": "Top Findings",
            "strategic_implications": "Strategic Implications",
            "rank": "Rank",
            "company_product": "Company/Product",
            "category": "Category(Direct/Adjacent/Substitute)",
            "target_user": "Target User",
            "core_jtbd": "Core JTBD",
            "platform": "Platform",
            "geo_focus": "Geo Focus",
            "traction_score": "Traction Score(1-5)",
            "product_capability_score": "Product Capability Score(1-5)",
            "monetization_score": "Monetization Score(1-5)",
            "user_sentiment_score": "User Sentiment Score(1-5)",
            "execution_maturity_score": "Execution Maturity Score(1-5)",
            "evidence_confidence_score": "Evidence Confidence Score(1-5)",
            "weighted_total": "Weighted Total(0-100)",
            "key_strength": "Key Strength",
            "key_weakness": "Key Weakness",
            "threat_level": "Threat Level",
            "l1_capability": "L1 Capability",
            "l2_module": "L2 Module",
            "l3_feature": "L3 Feature",
            "our_status": "Our Status(None/Planned/Live)",
            "competitor_coverage": "Competitor Coverage(0/1)",
            "parity_gap": "Parity Gap",
            "importance": "Importance(H/M/L)",
            "priority": "Priority",
            "product": "Product",
            "pricing_model": "Pricing Model",
            "entry_price": "Entry Price",
            "top_tier_price": "Top Tier Price",
            "trial_freemium": "Trial/Freemium",
            "packaging_unit": "Packaging Unit",
            "primary_channel": "Primary Channel(SEO/PLG/Sales/Partner)",
            "positioning_claim": "Positioning Claim",
            "observed_conversion_frictions": "Observed Conversion Frictions",
            "source_type": "Source Type(Official/Store/Review/Media/Research)",
            "url": "URL",
            "title": "Title",
            "published_date": "Published Date",
            "access_date": "Access Date",
            "claim": "Claim",
            "evidence_snippet": "Evidence Snippet",
            "confidence": "Confidence(High/Med/Low)",
        },
        "summary_templates": [
            {
                "problem_statement": "Market Definition",
                "target_segment": "Primary user segment to be validated",
                "method": "JTBD + competitor classification + weighted scoring",
                "top_findings": "Fill market definition based on brief/repo context",
                "strategic_implications": "Clarify category boundary before scoring",
            },
            {
                "problem_statement": "Positioning Statement",
                "target_segment": "Priority segment for GTM",
                "method": "Relative positioning vs Direct/Adjacent/Substitute",
                "top_findings": "Define what you compete on and what you intentionally do not",
                "strategic_implications": "Prevent scope creep and category drift",
            },
            {
                "problem_statement": "Strategic Implications",
                "target_segment": "Product + GTM stakeholders",
                "method": "Gap synthesis from benchmark and feature matrix",
                "top_findings": "Prioritize top 2-3 execution bets from competitor gaps",
                "strategic_implications": "Convert benchmark insights into roadmap decisions",
            },
        ],
        "enum": {
            "threat": {"high": "High", "medium": "Medium", "low": "Low"},
            "category": {
                "direct": "Direct",
                "adjacent": "Adjacent",
                "substitute": "Substitute",
            },
            "confidence": {"high": "High", "med": "Med", "low": "Low"},
            "our_status": {"none": "None", "planned": "Planned", "live": "Live"},
            "parity_gap": {
                "lead": "Lead",
                "parity": "Parity",
                "partial": "Partial",
                "gap": "Gap",
            },
        },
        "warnings": {
            "title": "Warnings:",
            "sources_lt3": "[WARN] {name}: sources < 3",
            "missing_official": "[WARN] {name}: missing official source",
            "missing_third": "[WARN] {name}: missing third-party source",
        },
    },
    "zh": {
        "sheet_names": {
            "summary": "摘要",
            "benchmark": "竞品基准",
            "feature_matrix": "功能矩阵",
            "pricing_gtm": "定价-GTM",
            "sources": "证据来源",
        },
        "headers": {
            "problem_statement": "问题定义",
            "target_segment": "目标用户",
            "method": "方法",
            "scope": "范围",
            "top_findings": "关键发现",
            "strategic_implications": "战略含义",
            "rank": "排名",
            "company_product": "公司/产品",
            "category": "分类(直接/邻近/替代)",
            "target_user": "目标人群",
            "core_jtbd": "核心JTBD",
            "platform": "平台",
            "geo_focus": "地域聚焦",
            "traction_score": "增长势能评分(1-5)",
            "product_capability_score": "产品能力评分(1-5)",
            "monetization_score": "商业化评分(1-5)",
            "user_sentiment_score": "用户口碑评分(1-5)",
            "execution_maturity_score": "执行成熟度评分(1-5)",
            "evidence_confidence_score": "证据可信度评分(1-5)",
            "weighted_total": "加权总分(0-100)",
            "key_strength": "主要优势",
            "key_weakness": "主要短板",
            "threat_level": "威胁等级",
            "l1_capability": "L1 能力域",
            "l2_module": "L2 模块",
            "l3_feature": "L3 功能",
            "our_status": "我方状态(None/Planned/Live)",
            "competitor_coverage": "竞品覆盖(0/1)",
            "parity_gap": "对位差距",
            "importance": "重要性(H/M/L)",
            "priority": "优先级",
            "product": "产品",
            "pricing_model": "定价模型",
            "entry_price": "入门价格",
            "top_tier_price": "最高档价格",
            "trial_freemium": "试用/免费",
            "packaging_unit": "计费单位",
            "primary_channel": "主要渠道(SEO/PLG/Sales/Partner)",
            "positioning_claim": "定位主张",
            "observed_conversion_frictions": "转化阻力观察",
            "source_type": "来源类型(官方/商店/评测/媒体/研究)",
            "url": "链接",
            "title": "标题",
            "published_date": "发布日期",
            "access_date": "访问日期",
            "claim": "结论主张",
            "evidence_snippet": "证据摘录",
            "confidence": "可信度(高/中/低)",
        },
        "summary_templates": [
            {
                "problem_statement": "市场定义",
                "target_segment": "待验证的核心用户群",
                "method": "JTBD + 竞品分层 + 加权评分",
                "top_findings": "基于需求描述或项目上下文补全市场定义",
                "strategic_implications": "先明确品类边界，再进入评分",
            },
            {
                "problem_statement": "定位陈述",
                "target_segment": "优先服务的GTM用户段",
                "method": "相对定位(直接/邻近/替代)",
                "top_findings": "定义“我们比什么”与“我们不比什么”",
                "strategic_implications": "防止范围蔓延与定位漂移",
            },
            {
                "problem_statement": "战略含义",
                "target_segment": "产品与增长决策团队",
                "method": "基准表 + 功能矩阵差距综合",
                "top_findings": "提炼2-3个最高优先级执行方向",
                "strategic_implications": "将竞品结论直接转为路线图决策",
            },
        ],
        "enum": {
            "threat": {"high": "高", "medium": "中", "low": "低"},
            "category": {
                "direct": "直接竞品",
                "adjacent": "邻近竞品",
                "substitute": "替代方案",
            },
            "confidence": {"high": "高", "med": "中", "low": "低"},
            "our_status": {"none": "无", "planned": "规划中", "live": "已上线"},
            "parity_gap": {
                "lead": "领先",
                "parity": "同等",
                "partial": "部分差距",
                "gap": "差距",
            },
        },
        "warnings": {
            "title": "警告:",
            "sources_lt3": "[WARN] {name}: 来源少于3条",
            "missing_official": "[WARN] {name}: 缺少官方来源",
            "missing_third": "[WARN] {name}: 缺少第三方来源",
        },
    },
    "ja": {
        "sheet_names": {
            "summary": "サマリー",
            "benchmark": "ベンチマーク",
            "feature_matrix": "機能マトリクス",
            "pricing_gtm": "価格・GTM",
            "sources": "ソース",
        },
        "headers": {
            "problem_statement": "課題定義",
            "target_segment": "ターゲットセグメント",
            "method": "手法",
            "scope": "範囲",
            "top_findings": "主要な発見",
            "strategic_implications": "戦略的示唆",
            "rank": "順位",
            "company_product": "企業/製品",
            "category": "分類(直接/隣接/代替)",
            "target_user": "対象ユーザー",
            "core_jtbd": "コアJTBD",
            "platform": "プラットフォーム",
            "geo_focus": "地域フォーカス",
            "traction_score": "トラクションスコア(1-5)",
            "product_capability_score": "製品能力スコア(1-5)",
            "monetization_score": "収益化スコア(1-5)",
            "user_sentiment_score": "ユーザー評価スコア(1-5)",
            "execution_maturity_score": "実行成熟度スコア(1-5)",
            "evidence_confidence_score": "証拠信頼度スコア(1-5)",
            "weighted_total": "加重合計(0-100)",
            "key_strength": "強み",
            "key_weakness": "弱み",
            "threat_level": "脅威レベル",
            "l1_capability": "L1 能力",
            "l2_module": "L2 モジュール",
            "l3_feature": "L3 機能",
            "our_status": "自社ステータス(None/Planned/Live)",
            "competitor_coverage": "競合カバー率(0/1)",
            "parity_gap": "ギャップ",
            "importance": "重要度(H/M/L)",
            "priority": "優先度",
            "product": "製品",
            "pricing_model": "価格モデル",
            "entry_price": "開始価格",
            "top_tier_price": "上位価格",
            "trial_freemium": "トライアル/無料",
            "packaging_unit": "課金単位",
            "primary_channel": "主要チャネル(SEO/PLG/Sales/Partner)",
            "positioning_claim": "ポジショニング",
            "observed_conversion_frictions": "転換障壁",
            "source_type": "ソース種別(公式/ストア/レビュー/メディア/調査)",
            "url": "URL",
            "title": "タイトル",
            "published_date": "公開日",
            "access_date": "アクセス日",
            "claim": "主張",
            "evidence_snippet": "証拠抜粋",
            "confidence": "信頼度(高/中/低)",
        },
        "summary_templates": [
            {
                "problem_statement": "市場定義",
                "target_segment": "検証対象の主要ユーザー層",
                "method": "JTBD + 競合分類 + 加重評価",
                "top_findings": "要件またはリポジトリ文脈から市場定義を補完",
                "strategic_implications": "評価前にカテゴリ境界を明確化",
            },
            {
                "problem_statement": "ポジショニング",
                "target_segment": "優先すべきGTMセグメント",
                "method": "直接/隣接/代替の相対比較",
                "top_findings": "何で勝負し、何で勝負しないかを定義",
                "strategic_implications": "スコープ拡散とポジションずれを防止",
            },
            {
                "problem_statement": "戦略的示唆",
                "target_segment": "プロダクト/成長チーム",
                "method": "ベンチマークと機能ギャップの統合",
                "top_findings": "優先度の高い実行テーマを2-3件抽出",
                "strategic_implications": "競合分析をロードマップ判断へ直結",
            },
        ],
        "enum": {
            "threat": {"high": "高", "medium": "中", "low": "低"},
            "category": {
                "direct": "直接競合",
                "adjacent": "隣接競合",
                "substitute": "代替",
            },
            "confidence": {"high": "高", "med": "中", "low": "低"},
            "our_status": {"none": "未対応", "planned": "計画中", "live": "提供中"},
            "parity_gap": {
                "lead": "優位",
                "parity": "同等",
                "partial": "部分ギャップ",
                "gap": "ギャップ",
            },
        },
        "warnings": {
            "title": "警告:",
            "sources_lt3": "[WARN] {name}: ソースが3件未満",
            "missing_official": "[WARN] {name}: 公式ソース不足",
            "missing_third": "[WARN] {name}: 第三者ソース不足",
        },
    },
    "ko": {
        "sheet_names": {
            "summary": "요약",
            "benchmark": "벤치마크",
            "feature_matrix": "기능 매트릭스",
            "pricing_gtm": "가격-GTM",
            "sources": "출처",
        },
        "headers": {
            "problem_statement": "문제 정의",
            "target_segment": "타깃 세그먼트",
            "method": "방법",
            "scope": "범위",
            "top_findings": "핵심 발견",
            "strategic_implications": "전략적 시사점",
            "rank": "순위",
            "company_product": "회사/제품",
            "category": "분류(직접/인접/대체)",
            "target_user": "타깃 사용자",
            "core_jtbd": "핵심 JTBD",
            "platform": "플랫폼",
            "geo_focus": "지역 포커스",
            "traction_score": "성장 점수(1-5)",
            "product_capability_score": "제품 역량 점수(1-5)",
            "monetization_score": "수익화 점수(1-5)",
            "user_sentiment_score": "사용자 평판 점수(1-5)",
            "execution_maturity_score": "실행 성숙도 점수(1-5)",
            "evidence_confidence_score": "근거 신뢰도 점수(1-5)",
            "weighted_total": "가중 총점(0-100)",
            "key_strength": "강점",
            "key_weakness": "약점",
            "threat_level": "위협 수준",
            "l1_capability": "L1 역량",
            "l2_module": "L2 모듈",
            "l3_feature": "L3 기능",
            "our_status": "우리 상태(None/Planned/Live)",
            "competitor_coverage": "경쟁사 커버리지(0/1)",
            "parity_gap": "격차",
            "importance": "중요도(H/M/L)",
            "priority": "우선순위",
            "product": "제품",
            "pricing_model": "가격 모델",
            "entry_price": "진입 가격",
            "top_tier_price": "상위 가격",
            "trial_freemium": "체험/무료",
            "packaging_unit": "과금 단위",
            "primary_channel": "주요 채널(SEO/PLG/Sales/Partner)",
            "positioning_claim": "포지셔닝",
            "observed_conversion_frictions": "전환 저해 요인",
            "source_type": "출처 유형(공식/스토어/리뷰/미디어/리서치)",
            "url": "URL",
            "title": "제목",
            "published_date": "게시일",
            "access_date": "접근일",
            "claim": "주장",
            "evidence_snippet": "근거 요약",
            "confidence": "신뢰도(높음/중간/낮음)",
        },
        "summary_templates": [
            {
                "problem_statement": "시장 정의",
                "target_segment": "검증 대상 핵심 사용자군",
                "method": "JTBD + 경쟁 분류 + 가중 점수",
                "top_findings": "요구사항/레포 문맥 기반 시장 정의 보완",
                "strategic_implications": "평가 전에 카테고리 경계를 명확히",
            },
            {
                "problem_statement": "포지셔닝 진술",
                "target_segment": "우선 공략 GTM 세그먼트",
                "method": "직접/인접/대체 상대 비교",
                "top_findings": "무엇으로 경쟁하고 무엇은 경쟁하지 않을지 정의",
                "strategic_implications": "범위 확장과 포지션 흔들림 방지",
            },
            {
                "problem_statement": "전략적 시사점",
                "target_segment": "제품 및 성장 의사결정자",
                "method": "벤치마크 + 기능 격차 통합",
                "top_findings": "상위 2~3개 실행 우선순위 도출",
                "strategic_implications": "경쟁 분석을 로드맵으로 연결",
            },
        ],
        "enum": {
            "threat": {"high": "높음", "medium": "중간", "low": "낮음"},
            "category": {
                "direct": "직접 경쟁",
                "adjacent": "인접 경쟁",
                "substitute": "대체재",
            },
            "confidence": {"high": "높음", "med": "중간", "low": "낮음"},
            "our_status": {"none": "없음", "planned": "계획", "live": "운영"},
            "parity_gap": {
                "lead": "우위",
                "parity": "동등",
                "partial": "부분 격차",
                "gap": "격차",
            },
        },
        "warnings": {
            "title": "경고:",
            "sources_lt3": "[WARN] {name}: 출처가 3개 미만",
            "missing_official": "[WARN] {name}: 공식 출처 누락",
            "missing_third": "[WARN] {name}: 제3자 출처 누락",
        },
    },
    "es": {
        "sheet_names": {
            "summary": "Resumen",
            "benchmark": "Benchmark",
            "feature_matrix": "Matriz de Funciones",
            "pricing_gtm": "Precios-GTM",
            "sources": "Fuentes",
        },
        "headers": {
            "problem_statement": "Definición del Problema",
            "target_segment": "Segmento Objetivo",
            "method": "Método",
            "scope": "Alcance",
            "top_findings": "Hallazgos Clave",
            "strategic_implications": "Implicaciones Estratégicas",
            "rank": "Ranking",
            "company_product": "Empresa/Producto",
            "category": "Categoría(Directo/Adyacente/Sustituto)",
            "target_user": "Usuario Objetivo",
            "core_jtbd": "JTBD Principal",
            "platform": "Plataforma",
            "geo_focus": "Foco Geográfico",
            "traction_score": "Score de Tracción(1-5)",
            "product_capability_score": "Score de Capacidad de Producto(1-5)",
            "monetization_score": "Score de Monetización(1-5)",
            "user_sentiment_score": "Score de Opinión de Usuario(1-5)",
            "execution_maturity_score": "Score de Madurez de Ejecución(1-5)",
            "evidence_confidence_score": "Score de Confianza de Evidencia(1-5)",
            "weighted_total": "Total Ponderado(0-100)",
            "key_strength": "Fortaleza Clave",
            "key_weakness": "Debilidad Clave",
            "threat_level": "Nivel de Amenaza",
            "l1_capability": "Capacidad L1",
            "l2_module": "Módulo L2",
            "l3_feature": "Función L3",
            "our_status": "Estado Nuestro(None/Planned/Live)",
            "competitor_coverage": "Cobertura de Competidor(0/1)",
            "parity_gap": "Brecha de Paridad",
            "importance": "Importancia(H/M/L)",
            "priority": "Prioridad",
            "product": "Producto",
            "pricing_model": "Modelo de Precio",
            "entry_price": "Precio de Entrada",
            "top_tier_price": "Precio Superior",
            "trial_freemium": "Prueba/Freemium",
            "packaging_unit": "Unidad de Paquete",
            "primary_channel": "Canal Principal(SEO/PLG/Sales/Partner)",
            "positioning_claim": "Propuesta de Posicionamiento",
            "observed_conversion_frictions": "Fricciones de Conversión Observadas",
            "source_type": "Tipo de Fuente(Official/Store/Review/Media/Research)",
            "url": "URL",
            "title": "Título",
            "published_date": "Fecha de Publicación",
            "access_date": "Fecha de Acceso",
            "claim": "Afirmación",
            "evidence_snippet": "Extracto de Evidencia",
            "confidence": "Confianza(Alta/Media/Baja)",
        },
        "summary_templates": [
            {
                "problem_statement": "Definición de Mercado",
                "target_segment": "Segmento principal por validar",
                "method": "JTBD + clasificación de competidores + scoring ponderado",
                "top_findings": "Completar definición de mercado según brief/contexto",
                "strategic_implications": "Aclarar frontera de categoría antes de puntuar",
            },
            {
                "problem_statement": "Declaración de Posicionamiento",
                "target_segment": "Segmento GTM prioritario",
                "method": "Posicionamiento relativo vs Directo/Adyacente/Sustituto",
                "top_findings": "Definir en qué competir y en qué no",
                "strategic_implications": "Evitar deriva de alcance y categoría",
            },
            {
                "problem_statement": "Implicaciones Estratégicas",
                "target_segment": "Stakeholders de producto y GTM",
                "method": "Síntesis de brechas desde benchmark y matriz de funciones",
                "top_findings": "Priorizar 2-3 apuestas de ejecución",
                "strategic_implications": "Convertir benchmark en decisiones de roadmap",
            },
        ],
        "enum": {
            "threat": {"high": "Alto", "medium": "Medio", "low": "Bajo"},
            "category": {
                "direct": "Directo",
                "adjacent": "Adyacente",
                "substitute": "Sustituto",
            },
            "confidence": {"high": "Alta", "med": "Media", "low": "Baja"},
            "our_status": {"none": "Ninguno", "planned": "Planificado", "live": "Activo"},
            "parity_gap": {
                "lead": "Lidera",
                "parity": "Paridad",
                "partial": "Parcial",
                "gap": "Brecha",
            },
        },
        "warnings": {
            "title": "Advertencias:",
            "sources_lt3": "[WARN] {name}: fuentes < 3",
            "missing_official": "[WARN] {name}: falta fuente oficial",
            "missing_third": "[WARN] {name}: falta fuente de terceros",
        },
    },
    "fr": {
        "sheet_names": {
            "summary": "Résumé",
            "benchmark": "Benchmark",
            "feature_matrix": "Matrice Fonctionnelle",
            "pricing_gtm": "Prix-GTM",
            "sources": "Sources",
        },
        "headers": {
            "problem_statement": "Définition du Problème",
            "target_segment": "Segment Cible",
            "method": "Méthode",
            "scope": "Périmètre",
            "top_findings": "Principaux Résultats",
            "strategic_implications": "Implications Stratégiques",
            "rank": "Rang",
            "company_product": "Entreprise/Produit",
            "category": "Catégorie(Direct/Adjacent/Substitut)",
            "target_user": "Utilisateur Cible",
            "core_jtbd": "JTBD Central",
            "platform": "Plateforme",
            "geo_focus": "Cible Géographique",
            "traction_score": "Score de Traction(1-5)",
            "product_capability_score": "Score Capacité Produit(1-5)",
            "monetization_score": "Score Monétisation(1-5)",
            "user_sentiment_score": "Score Sentiment Utilisateur(1-5)",
            "execution_maturity_score": "Score Maturité Exécution(1-5)",
            "evidence_confidence_score": "Score Confiance Preuve(1-5)",
            "weighted_total": "Total Pondéré(0-100)",
            "key_strength": "Force Clé",
            "key_weakness": "Faiblesse Clé",
            "threat_level": "Niveau de Menace",
            "l1_capability": "Capacité L1",
            "l2_module": "Module L2",
            "l3_feature": "Fonctionnalité L3",
            "our_status": "Notre Statut(None/Planned/Live)",
            "competitor_coverage": "Couverture Concurrent(0/1)",
            "parity_gap": "Écart de Parité",
            "importance": "Importance(H/M/L)",
            "priority": "Priorité",
            "product": "Produit",
            "pricing_model": "Modèle Tarifaire",
            "entry_price": "Prix d'Entrée",
            "top_tier_price": "Prix Max",
            "trial_freemium": "Essai/Freemium",
            "packaging_unit": "Unité de Packaging",
            "primary_channel": "Canal Principal(SEO/PLG/Sales/Partner)",
            "positioning_claim": "Promesse de Positionnement",
            "observed_conversion_frictions": "Friction de Conversion Observée",
            "source_type": "Type de Source(Official/Store/Review/Media/Research)",
            "url": "URL",
            "title": "Titre",
            "published_date": "Date de Publication",
            "access_date": "Date d'Accès",
            "claim": "Assertion",
            "evidence_snippet": "Extrait de Preuve",
            "confidence": "Confiance(Haut/Moyen/Bas)",
        },
        "summary_templates": [
            {
                "problem_statement": "Définition du Marché",
                "target_segment": "Segment principal à valider",
                "method": "JTBD + classification concurrentielle + score pondéré",
                "top_findings": "Compléter selon brief/contexte projet",
                "strategic_implications": "Clarifier la frontière de catégorie avant scoring",
            },
            {
                "problem_statement": "Positionnement",
                "target_segment": "Segment GTM prioritaire",
                "method": "Positionnement relatif vs Direct/Adjacent/Substitut",
                "top_findings": "Définir ce qui est comparé et ce qui ne l'est pas",
                "strategic_implications": "Éviter dérive de périmètre et de catégorie",
            },
            {
                "problem_statement": "Implications Stratégiques",
                "target_segment": "Parties prenantes produit et GTM",
                "method": "Synthèse des écarts benchmark + matrice fonctionnelle",
                "top_findings": "Prioriser 2-3 paris d'exécution",
                "strategic_implications": "Transformer le benchmark en décisions roadmap",
            },
        ],
        "enum": {
            "threat": {"high": "Élevé", "medium": "Moyen", "low": "Faible"},
            "category": {
                "direct": "Direct",
                "adjacent": "Adjacent",
                "substitute": "Substitut",
            },
            "confidence": {"high": "Haut", "med": "Moyen", "low": "Bas"},
            "our_status": {"none": "Aucun", "planned": "Planifié", "live": "En ligne"},
            "parity_gap": {
                "lead": "Avance",
                "parity": "Parité",
                "partial": "Partiel",
                "gap": "Écart",
            },
        },
        "warnings": {
            "title": "Avertissements:",
            "sources_lt3": "[WARN] {name}: sources < 3",
            "missing_official": "[WARN] {name}: source officielle manquante",
            "missing_third": "[WARN] {name}: source tierce manquante",
        },
    },
    "de": {
        "sheet_names": {
            "summary": "Zusammenfassung",
            "benchmark": "Benchmark",
            "feature_matrix": "Feature-Matrix",
            "pricing_gtm": "Pricing-GTM",
            "sources": "Quellen",
        },
        "headers": {
            "problem_statement": "Problemdefinition",
            "target_segment": "Zielsegment",
            "method": "Methode",
            "scope": "Umfang",
            "top_findings": "Wichtigste Erkenntnisse",
            "strategic_implications": "Strategische Implikationen",
            "rank": "Rang",
            "company_product": "Unternehmen/Produkt",
            "category": "Kategorie(Direkt/Angrenzend/Ersatz)",
            "target_user": "Zielnutzer",
            "core_jtbd": "Kern-JTBD",
            "platform": "Plattform",
            "geo_focus": "Geo-Fokus",
            "traction_score": "Traction-Score(1-5)",
            "product_capability_score": "Produktfähigkeits-Score(1-5)",
            "monetization_score": "Monetarisierungs-Score(1-5)",
            "user_sentiment_score": "Nutzerstimmungs-Score(1-5)",
            "execution_maturity_score": "Reifegrad-Score(1-5)",
            "evidence_confidence_score": "Evidenz-Score(1-5)",
            "weighted_total": "Gewichtete Summe(0-100)",
            "key_strength": "Stärke",
            "key_weakness": "Schwäche",
            "threat_level": "Bedrohungsgrad",
            "l1_capability": "L1 Fähigkeit",
            "l2_module": "L2 Modul",
            "l3_feature": "L3 Feature",
            "our_status": "Unser Status(None/Planned/Live)",
            "competitor_coverage": "Wettbewerbsabdeckung(0/1)",
            "parity_gap": "Paritätslücke",
            "importance": "Wichtigkeit(H/M/L)",
            "priority": "Priorität",
            "product": "Produkt",
            "pricing_model": "Preismodell",
            "entry_price": "Einstiegspreis",
            "top_tier_price": "Top-Preis",
            "trial_freemium": "Test/Freemium",
            "packaging_unit": "Paketeinheit",
            "primary_channel": "Hauptkanal(SEO/PLG/Sales/Partner)",
            "positioning_claim": "Positionierungsversprechen",
            "observed_conversion_frictions": "Beobachtete Conversion-Hürden",
            "source_type": "Quellentyp(Official/Store/Review/Media/Research)",
            "url": "URL",
            "title": "Titel",
            "published_date": "Veröffentlichungsdatum",
            "access_date": "Abrufdatum",
            "claim": "Aussage",
            "evidence_snippet": "Evidenz-Auszug",
            "confidence": "Vertrauen(Hoch/Mittel/Niedrig)",
        },
        "summary_templates": [
            {
                "problem_statement": "Marktdefinition",
                "target_segment": "Zu validierendes Kernsegment",
                "method": "JTBD + Wettbewerbscluster + gewichtetes Scoring",
                "top_findings": "Marktdefinition aus Brief/Repo-Kontext ergänzen",
                "strategic_implications": "Kategoriegrenze vor Scoring klären",
            },
            {
                "problem_statement": "Positionierungsstatement",
                "target_segment": "Priorisiertes GTM-Segment",
                "method": "Relative Positionierung vs Direkt/Angrenzend/Ersatz",
                "top_findings": "Definieren, worin konkurriert wird und worin nicht",
                "strategic_implications": "Scope- und Kategorie-Drift vermeiden",
            },
            {
                "problem_statement": "Strategische Implikationen",
                "target_segment": "Produkt- und GTM-Teams",
                "method": "Gap-Synthese aus Benchmark und Feature-Matrix",
                "top_findings": "Top 2-3 Umsetzungswetten priorisieren",
                "strategic_implications": "Benchmark in Roadmap-Entscheidungen überführen",
            },
        ],
        "enum": {
            "threat": {"high": "Hoch", "medium": "Mittel", "low": "Niedrig"},
            "category": {
                "direct": "Direkt",
                "adjacent": "Angrenzend",
                "substitute": "Ersatz",
            },
            "confidence": {"high": "Hoch", "med": "Mittel", "low": "Niedrig"},
            "our_status": {"none": "Kein", "planned": "Geplant", "live": "Live"},
            "parity_gap": {
                "lead": "Vorsprung",
                "parity": "Parität",
                "partial": "Teilweise",
                "gap": "Lücke",
            },
        },
        "warnings": {
            "title": "Warnungen:",
            "sources_lt3": "[WARN] {name}: Quellen < 3",
            "missing_official": "[WARN] {name}: offizielle Quelle fehlt",
            "missing_third": "[WARN] {name}: Drittquelle fehlt",
        },
    },
}

# Canonical tokens for multilingual normalization
ENUM_CANONICAL: dict[str, dict[str, list[str]]] = {
    "threat": {
        "high": ["high", "高", "높음", "alto", "élevé", "hoch"],
        "medium": ["medium", "med", "中", "중간", "medio", "moyen", "mittel"],
        "low": ["low", "低", "낮음", "bajo", "faible", "niedrig"],
    },
    "category": {
        "direct": ["direct", "直接", "直接競合", "직접", "directo", "direkt"],
        "adjacent": ["adjacent", "邻近", "隣接", "인접", "adyacente", "angrenzend"],
        "substitute": ["substitute", "替代", "代替", "대체", "sustituto", "ersatz", "substitut"],
    },
    "confidence": {
        "high": ["high", "高", "높음", "alta", "haut", "hoch"],
        "med": ["med", "medium", "中", "중간", "media", "moyen", "mittel"],
        "low": ["low", "低", "낮음", "baja", "bas", "niedrig"],
    },
    "our_status": {
        "none": ["none", "无", "未対応", "없음", "ninguno", "aucun", "kein"],
        "planned": ["planned", "规划", "計画", "계획", "planificado", "planifié", "geplant"],
        "live": ["live", "已上线", "提供中", "운영", "activo", "enligne", "live"],
    },
    "parity_gap": {
        "lead": ["lead", "领先", "優位", "우위", "lidera", "avance", "vorsprung"],
        "parity": ["parity", "同等", "同等", "동등", "paridad", "parité", "parität"],
        "partial": ["partial", "部分", "部分", "부분", "parcial", "partiel", "teilweise"],
        "gap": ["gap", "差距", "ギャップ", "격차", "brecha", "écart", "lücke"],
    },
}

SOURCE_TYPE_TOKENS = {
    "official": ["official", "官网", "官方", "公式", "공식"],
    "store": ["store", "appstore", "play", "商店", "스토어"],
    "review": ["review", "评测", "レビュー", "리뷰"],
    "media": ["media", "媒体", "メディア", "미디어"],
    "research": ["research", "报告", "研究", "リサーチ", "연구"],
}

LANG_TOKEN_HINTS = {
    "es": {"de", "la", "para", "con", "que", "los", "las", "una", "por"},
    "fr": {"le", "la", "les", "de", "des", "et", "pour", "avec", "une"},
    "de": {"und", "der", "die", "das", "mit", "für", "ein", "eine", "nicht"},
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build competitive benchmark XLSX")
    parser.add_argument("--output", required=True, type=Path, help="Output XLSX path")
    parser.add_argument("--input-json", type=Path, help="Structured input JSON path")
    parser.add_argument("--brief", default="", help="Requirement brief")
    parser.add_argument("--project-path", default="", help="Project path for context")
    parser.add_argument("--region", default="global", help="Target region")
    parser.add_argument("--top-n", type=int, default=8, help="Competitor count cap")
    parser.add_argument(
        "--period-months", type=int, default=24, help="Evidence lookback window"
    )
    parser.add_argument(
        "--weights",
        default="20,30,15,20,10,5",
        help="Weights for Traction,Capability,Monetization,Sentiment,Execution,Evidence",
    )
    parser.add_argument(
        "--lang",
        default="auto",
        choices=["auto", "en", "zh", "ja", "ko", "es", "fr", "de"],
        help="Output language. Default: auto detect from input text",
    )
    parser.add_argument(
        "--lang-source",
        default="both",
        choices=["brief", "input", "both"],
        help="Text source used for automatic language detection",
    )
    return parser.parse_args()


def norm_key(text: str) -> str:
    return "".join(ch.lower() for ch in str(text) if ch.isalnum())


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_weights(raw: str) -> list[float]:
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) != 6:
        raise ValueError("--weights must contain exactly 6 numbers")
    weights = [float(p) for p in parts]
    if abs(sum(weights) - 100.0) > 1e-6:
        raise ValueError("weights must sum to 100")
    return weights


def col_to_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def load_payload(path: Path | None) -> dict[str, Any]:
    if not path:
        return {}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def get_locale(lang: str) -> dict[str, Any]:
    if lang in LOCALES:
        return LOCALES[lang]
    return LOCALES["en"]


def build_scope_text(args: argparse.Namespace) -> str:
    scope_parts = []
    if args.project_path:
        scope_parts.append(f"project={args.project_path}")
    if args.brief:
        scope_parts.append("brief provided")
    scope_parts.append(f"region={args.region}")
    scope_parts.append(f"top_n={args.top_n}")
    scope_parts.append(f"window={args.period_months}m")
    return "; ".join(scope_parts)


def build_default_summary(args: argparse.Namespace, lang: str) -> list[dict[str, Any]]:
    locale = get_locale(lang)
    templates = locale["summary_templates"]
    scope_text = build_scope_text(args)

    rows: list[dict[str, Any]] = []
    for idx, tpl in enumerate(templates):
        row = {
            "problem_statement": tpl["problem_statement"],
            "target_segment": tpl["target_segment"],
            "method": tpl["method"],
            "scope": scope_text,
            "top_findings": tpl["top_findings"],
            "strategic_implications": tpl["strategic_implications"],
        }
        if idx == 0 and args.brief:
            row["top_findings"] = args.brief
        rows.append(row)
    return rows


def extract_detection_text(
    args: argparse.Namespace, payload: dict[str, Any]
) -> str:
    parts: list[str] = []
    if args.lang_source in ("brief", "both") and args.brief.strip():
        parts.append(args.brief)
    if args.lang_source in ("input", "both") and payload:
        parts.append(json.dumps(payload, ensure_ascii=False))
    return "\n".join(parts)


def detect_language(text: str) -> str:
    if not text.strip():
        return "en"

    count_han = 0
    count_kana = 0
    count_hangul = 0
    count_latin = 0

    for ch in text:
        code = ord(ch)
        if 0x3040 <= code <= 0x30FF:
            count_kana += 1
        elif 0xAC00 <= code <= 0xD7AF:
            count_hangul += 1
        elif 0x4E00 <= code <= 0x9FFF:
            count_han += 1
        elif ("a" <= ch.lower() <= "z"):
            count_latin += 1

    major_scripts = sorted(
        [
            ("han", count_han),
            ("kana", count_kana),
            ("hangul", count_hangul),
            ("latin", count_latin),
        ],
        key=lambda x: x[1],
        reverse=True,
    )
    top_name, top_count = major_scripts[0]
    second_count = major_scripts[1][1]
    total = sum(v for _, v in major_scripts)

    if total > 0 and abs(top_count - second_count) / total < 0.10:
        return "en"

    if count_kana >= 6:
        return "ja"
    if count_hangul >= 6:
        return "ko"
    if count_han >= 10 and count_kana < 6 and count_hangul < 6:
        return "zh"

    if count_latin == 0:
        return "en"

    tokens = re.findall(r"[a-zA-ZÀ-ÿ]+", text.lower())
    lang_scores = {"es": 0, "fr": 0, "de": 0}
    for t in tokens:
        for lang, hints in LANG_TOKEN_HINTS.items():
            if t in hints:
                lang_scores[lang] += 1

    best_lang = max(lang_scores, key=lang_scores.get)
    if lang_scores[best_lang] >= 2:
        return best_lang
    return "en"


def choose_language(args: argparse.Namespace, payload: dict[str, Any]) -> str:
    if args.lang != "auto":
        return args.lang
    detected = detect_language(extract_detection_text(args, payload))
    if detected in LOCALES:
        return detected
    return "en"


def sheet_header(lang: str, col_key: str) -> str:
    locale = get_locale(lang)
    return locale["headers"].get(col_key, LOCALES["en"]["headers"].get(col_key, col_key))


def sheet_name(lang: str, sheet_key: str) -> str:
    locale = get_locale(lang)
    return locale["sheet_names"].get(sheet_key, LOCALES["en"]["sheet_names"][sheet_key])


def build_sheet_aliases() -> dict[str, list[str]]:
    aliases: dict[str, set[str]] = {k: set() for k in SHEET_ORDER}
    for key in SHEET_ORDER:
        aliases[key].update({key, key.replace("_", "-"), key.replace("_", "")})
        for loc in LOCALES.values():
            aliases[key].add(loc["sheet_names"][key])
    return {k: [norm_key(x) for x in v if x] for k, v in aliases.items()}


def canonical_sheet_key(raw_key: str, sheet_aliases: dict[str, list[str]]) -> str | None:
    nk = norm_key(raw_key)
    for key, values in sheet_aliases.items():
        if nk in values:
            return key
    return None


def get_payload_rows(
    payload: dict[str, Any], sheet_key: str, sheet_aliases: dict[str, list[str]]
) -> list[dict[str, Any]]:
    for k, v in payload.items():
        canonical = canonical_sheet_key(k, sheet_aliases)
        if canonical == sheet_key and isinstance(v, list):
            return [x for x in v if isinstance(x, dict)]
    return []


def build_column_aliases() -> dict[str, dict[str, list[str]]]:
    aliases: dict[str, dict[str, set[str]]] = {
        sheet: {col: {col} for col in conf["columns"]}
        for sheet, conf in SHEET_LAYOUT.items()
    }

    for sheet_key, conf in SHEET_LAYOUT.items():
        for col in conf["columns"]:
            aliases[sheet_key][col].add(LOCALES["en"]["headers"][col])

    for locale in LOCALES.values():
        for sheet_key, conf in SHEET_LAYOUT.items():
            for col in conf["columns"]:
                aliases[sheet_key][col].add(locale["headers"][col])

    return {
        s: {c: [norm_key(x) for x in vals if x] for c, vals in d.items()}
        for s, d in aliases.items()
    }


def map_rows(
    raw_rows: list[dict[str, Any]],
    sheet_key: str,
    column_aliases: dict[str, dict[str, list[str]]],
) -> list[dict[str, Any]]:
    mapped: list[dict[str, Any]] = []
    columns = SHEET_LAYOUT[sheet_key]["columns"]

    for row in raw_rows:
        key_map = {norm_key(k): v for k, v in row.items()}
        out: dict[str, Any] = {}
        for col in columns:
            out[col] = ""
            for alias in column_aliases[sheet_key][col]:
                if alias in key_map:
                    out[col] = key_map[alias]
                    break
        mapped.append(out)
    return mapped


def canonical_from_value(kind: str, value: Any) -> str | None:
    if value is None:
        return None
    raw = norm_key(str(value))
    if not raw:
        return None
    for canonical, variants in ENUM_CANONICAL.get(kind, {}).items():
        for token in variants:
            if norm_key(token) == raw:
                return canonical
    return None


def localize_enum(kind: str, value: Any, lang: str) -> Any:
    if value is None or value == "":
        return value
    canonical = canonical_from_value(kind, value)
    if canonical is None:
        return value
    locale = get_locale(lang)
    localized = locale["enum"].get(kind, {}).get(canonical)
    if localized:
        return localized
    return value


def normalize_source_type(value: Any) -> set[str]:
    if value is None:
        return set()
    s = norm_key(str(value))
    found: set[str] = set()
    for k, tokens in SOURCE_TYPE_TOKENS.items():
        for t in tokens:
            if norm_key(t) in s:
                found.add(k)
                break
    return found


def localize_rows_for_output(rows: list[dict[str, Any]], sheet_key: str, lang: str) -> None:
    if sheet_key == "benchmark":
        for row in rows:
            row["category"] = localize_enum("category", row.get("category"), lang)
            row["threat_level"] = localize_enum("threat", row.get("threat_level"), lang)
    elif sheet_key == "sources":
        for row in rows:
            row["confidence"] = localize_enum("confidence", row.get("confidence"), lang)
    elif sheet_key == "feature_matrix":
        for row in rows:
            row["our_status"] = localize_enum("our_status", row.get("our_status"), lang)
            row["parity_gap"] = localize_enum("parity_gap", row.get("parity_gap"), lang)


def prepare_benchmark_rows(
    rows: list[dict[str, Any]], weights: list[float], top_n: int, lang: str
) -> list[dict[str, Any]]:
    enriched: list[dict[str, Any]] = []
    for row in rows:
        scores = [
            to_float(row.get("traction_score")),
            to_float(row.get("product_capability_score")),
            to_float(row.get("monetization_score")),
            to_float(row.get("user_sentiment_score")),
            to_float(row.get("execution_maturity_score")),
            to_float(row.get("evidence_confidence_score")),
        ]
        if all(v is not None for v in scores):
            weighted = sum((scores[i] / 5.0) * weights[i] for i in range(6))
            row["__weighted_value"] = round(weighted, 2)
        else:
            row["__weighted_value"] = None
        enriched.append(row)

    enriched.sort(
        key=lambda r: r["__weighted_value"] if r["__weighted_value"] is not None else -1,
        reverse=True,
    )
    enriched = enriched[:top_n]

    for idx, row in enumerate(enriched, start=1):
        row["rank"] = idx
        if row["__weighted_value"] is None:
            row["weighted_total"] = ""
        if not str(row.get("threat_level", "")).strip() and row["__weighted_value"] is not None:
            wv = row["__weighted_value"]
            if wv >= 75:
                row["threat_level"] = localize_enum("threat", "high", lang)
            elif wv >= 55:
                row["threat_level"] = localize_enum("threat", "medium", lang)
            else:
                row["threat_level"] = localize_enum("threat", "low", lang)
    return enriched


def validate_sources(rows: list[dict[str, Any]], lang: str) -> list[str]:
    locale = get_locale(lang)
    warnings = locale["warnings"]

    bucket: dict[str, list[set[str]]] = {}
    for row in rows:
        name = str(row.get("product", "")).strip()
        if not name:
            continue
        bucket.setdefault(name, []).append(normalize_source_type(row.get("source_type")))

    out: list[str] = []
    for name, types_list in bucket.items():
        merged: set[str] = set()
        for t in types_list:
            merged.update(t)

        if len(types_list) < 3:
            out.append(warnings["sources_lt3"].format(name=name))
        if "official" not in merged:
            out.append(warnings["missing_official"].format(name=name))
        if not merged.intersection({"store", "review", "media", "research"}):
            out.append(warnings["missing_third"].format(name=name))
    return out


def worksheet_xml(
    rows: list[list[Any]], widths: list[float], freeze_header: bool = True
) -> str:
    max_cols = max(1, max((len(r) for r in rows), default=1))
    total_rows = max(1, len(rows))

    cols_xml = []
    for idx in range(1, max_cols + 1):
        w = widths[idx - 1] if idx - 1 < len(widths) else 24.0
        cols_xml.append(f'<col min="{idx}" max="{idx}" width="{w}" customWidth="1"/>')

    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    if freeze_header:
        lines.append(
            '  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        )
    else:
        lines.append('  <sheetViews><sheetView workbookViewId="0"/></sheetViews>')

    lines.extend(
        [
            '  <sheetFormatPr defaultRowHeight="22"/>',
            f'  <cols>{"".join(cols_xml)}</cols>',
            f'  <dimension ref="A1:{col_to_letter(max_cols)}{total_rows}"/>',
            '  <sheetData>',
        ]
    )

    for r_idx, row in enumerate(rows, start=1):
        lines.append(f'    <row r="{r_idx}" ht="22" customHeight="1">')
        for c_idx in range(1, max_cols + 1):
            val = row[c_idx - 1] if c_idx - 1 < len(row) else ""
            ref = f"{col_to_letter(c_idx)}{r_idx}"
            style = "2" if r_idx == 1 else "1"

            if val is None or val == "":
                lines.append(f'      <c r="{ref}" s="{style}"/>')
            elif isinstance(val, (int, float)):
                lines.append(f'      <c r="{ref}" s="{style}"><v>{val}</v></c>')
            elif isinstance(val, str) and val.startswith("="):
                formula = escape(val[1:])
                lines.append(f'      <c r="{ref}" s="{style}"><f>{formula}</f></c>')
            else:
                txt = str(val)
                esc = escape(txt)
                preserve = (
                    ' xml:space="preserve"'
                    if txt.startswith(" ") or txt.endswith(" ") or "\n" in txt
                    else ""
                )
                lines.append(
                    f'      <c r="{ref}" t="inlineStr" s="{style}"><is><t{preserve}>{esc}</t></is></c>'
                )
        lines.append("    </row>")

    lines.append("  </sheetData>")
    lines.append(f'  <autoFilter ref="A1:{col_to_letter(max_cols)}{total_rows}"/>')
    lines.append("</worksheet>")
    return "\n".join(lines)


def styles_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <fonts count=\"2\">
    <font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>
    <font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>
  </fonts>
  <fills count=\"2\">
    <fill><patternFill patternType=\"none\"/></fill>
    <fill><patternFill patternType=\"gray125\"/></fill>
  </fills>
  <borders count=\"1\">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count=\"1\">
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>
  </cellStyleXfs>
  <cellXfs count=\"3\">
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>
    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyFont=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>
  </cellXfs>
  <cellStyles count=\"1\">
    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>
  </cellStyles>
</styleSheet>
"""


def build_workbook_xml(sheet_names: list[str]) -> str:
    sheet_lines = []
    for idx, name in enumerate(sheet_names, start=1):
        safe_name = escape(name)
        sheet_lines.append(
            f'    <sheet name="{safe_name}" sheetId="{idx}" r:id="rId{idx}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        '  <sheets>\n'
        + "\n".join(sheet_lines)
        + '\n  </sheets>\n'
        '</workbook>\n'
    )


def build_workbook_rels(sheet_count: int) -> str:
    rels = []
    for idx in range(1, sheet_count + 1):
        rels.append(
            f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        )
    rels.append(
        f'  <Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        + "\n".join(rels)
        + '\n</Relationships>\n'
    )


def build_content_types(sheet_count: int) -> str:
    overrides = [
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for idx in range(1, sheet_count + 1):
        overrides.append(
            f'  <Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        + "\n".join(overrides)
        + '\n</Types>\n'
    )


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def core_xml() -> str:
    today = date.today().isoformat()
    ts = f"{today}T00:00:00Z"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>competitive-analysis-skill</dc:creator>
  <cp:lastModifiedBy>competitive-analysis-skill</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{ts}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{ts}</dcterms:modified>
</cp:coreProperties>
"""


def app_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>competitive-analysis-skill</Application>
</Properties>
"""


def write_xlsx(output: Path, sheet_entries: list[tuple[str, str]]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet_names = [name for name, _ in sheet_entries]

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types(len(sheet_names)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", build_workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(sheet_names)))
        zf.writestr("xl/styles.xml", styles_xml())
        for idx, (_, xml) in enumerate(sheet_entries, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml)
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml())


def to_sheet_rows(
    sheet_key: str,
    rows: list[dict[str, Any]],
    lang: str,
    weights: list[float] | None = None,
) -> list[list[Any]]:
    columns = SHEET_LAYOUT[sheet_key]["columns"]
    header = [sheet_header(lang, c) for c in columns]
    data_rows: list[list[Any]] = [header]

    if sheet_key != "benchmark":
        localize_rows_for_output(rows, sheet_key, lang)
        for item in rows:
            data_rows.append([item.get(c, "") for c in columns])
        return data_rows

    # Benchmark keeps formula for weighted_total
    assert weights is not None
    localize_rows_for_output(rows, sheet_key, lang)
    for idx, item in enumerate(rows, start=2):
        line: list[Any] = []
        for c in columns:
            if c == "weighted_total":
                scores = [
                    to_float(item.get("traction_score")),
                    to_float(item.get("product_capability_score")),
                    to_float(item.get("monetization_score")),
                    to_float(item.get("user_sentiment_score")),
                    to_float(item.get("execution_maturity_score")),
                    to_float(item.get("evidence_confidence_score")),
                ]
                if all(v is not None for v in scores):
                    formula = (
                        f"=ROUND(H{idx}/5*{weights[0]}+I{idx}/5*{weights[1]}+"
                        f"J{idx}/5*{weights[2]}+K{idx}/5*{weights[3]}+"
                        f"L{idx}/5*{weights[4]}+M{idx}/5*{weights[5]},2)"
                    )
                    line.append(formula)
                else:
                    line.append(item.get(c, ""))
            else:
                line.append(item.get(c, ""))
        data_rows.append(line)

    return data_rows


def main() -> None:
    args = parse_args()
    weights = parse_weights(args.weights)
    payload = load_payload(args.input_json)

    selected_lang = choose_language(args, payload)

    sheet_aliases = build_sheet_aliases()
    column_aliases = build_column_aliases()

    summary_raw = get_payload_rows(payload, "summary", sheet_aliases)
    if summary_raw:
        summary_rows_data = map_rows(summary_raw, "summary", column_aliases)
    else:
        summary_rows_data = build_default_summary(args, selected_lang)

    benchmark_raw = get_payload_rows(payload, "benchmark", sheet_aliases)
    benchmark_rows_data = map_rows(benchmark_raw, "benchmark", column_aliases)
    benchmark_rows_data = prepare_benchmark_rows(
        benchmark_rows_data, weights, args.top_n, selected_lang
    )

    feature_raw = get_payload_rows(payload, "feature_matrix", sheet_aliases)
    feature_rows_data = map_rows(feature_raw, "feature_matrix", column_aliases)

    pricing_raw = get_payload_rows(payload, "pricing_gtm", sheet_aliases)
    pricing_rows_data = map_rows(pricing_raw, "pricing_gtm", column_aliases)

    sources_raw = get_payload_rows(payload, "sources", sheet_aliases)
    sources_rows_data = map_rows(sources_raw, "sources", column_aliases)

    warnings = validate_sources(sources_rows_data, selected_lang)

    sheet_entries: list[tuple[str, str]] = []

    summary_rows = to_sheet_rows("summary", summary_rows_data, selected_lang)
    sheet_entries.append(
        (sheet_name(selected_lang, "summary"), worksheet_xml(summary_rows, SHEET_LAYOUT["summary"]["widths"]))
    )

    benchmark_rows = to_sheet_rows(
        "benchmark", benchmark_rows_data, selected_lang, weights=weights
    )
    sheet_entries.append(
        (
            sheet_name(selected_lang, "benchmark"),
            worksheet_xml(benchmark_rows, SHEET_LAYOUT["benchmark"]["widths"]),
        )
    )

    feature_rows = to_sheet_rows("feature_matrix", feature_rows_data, selected_lang)
    sheet_entries.append(
        (
            sheet_name(selected_lang, "feature_matrix"),
            worksheet_xml(feature_rows, SHEET_LAYOUT["feature_matrix"]["widths"]),
        )
    )

    pricing_rows = to_sheet_rows("pricing_gtm", pricing_rows_data, selected_lang)
    sheet_entries.append(
        (
            sheet_name(selected_lang, "pricing_gtm"),
            worksheet_xml(pricing_rows, SHEET_LAYOUT["pricing_gtm"]["widths"]),
        )
    )

    sources_rows = to_sheet_rows("sources", sources_rows_data, selected_lang)
    sheet_entries.append(
        (
            sheet_name(selected_lang, "sources"),
            worksheet_xml(sources_rows, SHEET_LAYOUT["sources"]["widths"]),
        )
    )

    write_xlsx(args.output, sheet_entries)

    print(f"Written: {args.output}")
    print(f"Language: {selected_lang}")
    print("Sheets:", ", ".join(name for name, _ in sheet_entries))
    print(f"Benchmark rows: {max(len(benchmark_rows) - 1, 0)}")

    if warnings:
        warn_title = get_locale(selected_lang)["warnings"]["title"]
        print(warn_title)
        for msg in warnings:
            print(msg)


if __name__ == "__main__":
    main()
