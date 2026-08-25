from decimal import Decimal

import pandas as pd

from hopptx import processor


def test_evaluation_count_uses_sessions_and_rounds(monkeypatch) -> None:
    monkeypatch.setattr(processor, "load_category_map", lambda _path: {})
    monkeypatch.setattr(
        processor,
        "categorize_topics",
        lambda topics, _category_map, _threshold: {"Uncategorized": topics},
    )

    df = pd.DataFrame(
        [
            {
                "Start Date & Time": "2026-01-01 10:00",
                "Event Created Date & Time": "2025-12-31 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "a@example.com",
            },
            {
                "Start Date & Time": "2026-01-02 10:00",
                "Event Created Date & Time": "2026-01-01 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "b@example.com",
            },
            {
                "Start Date & Time": "2026-01-03 10:00",
                "Event Created Date & Time": "2026-01-02 09:00",
                "Canceled": False,
                "Event Type Name": "Topic B",
                "Invitee Email": "c@example.com",
            },
            {
                "Start Date & Time": "2026-01-04 10:00",
                "Event Created Date & Time": "2026-01-03 09:00",
                "Canceled": True,
                "Event Type Name": "Topic B",
                "Invitee Email": "d@example.com",
            },
        ]
    )

    metrics = processor.compute_metrics(
        df=df,
        start="2026-01-01",
        end="2026-01-31",
        exclude_canceled=True,
        cost_per_training=Decimal("495"),
        evaluation_rate_percentage=Decimal("40"),
        top_topics_count=3,
        top_topics_hard_cutoff=10,
        group_topics=False,
        clean_topic_names=False,
        topic_fuzzy_match_threshold=90,
        category_reference="unused.docx",
        category_fuzzy_match_threshold=86,
    )

    assert metrics.total_sign_ups == 4
    assert metrics.total_attended_sign_ups == 3
    assert metrics.unique_topics == 2
    assert metrics.total_evaluation_count == 1


def test_evaluation_count_uses_unique_sessions(monkeypatch) -> None:
    monkeypatch.setattr(processor, "load_category_map", lambda _path: {})
    monkeypatch.setattr(
        processor,
        "categorize_topics",
        lambda topics, _category_map, _threshold: {"Uncategorized": topics},
    )

    df = pd.DataFrame(
        [
            {
                "Start Date & Time": "2026-03-01 10:00",
                "Event Created Date & Time": "2026-02-15 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "a@example.com",
            },
            {
                "Start Date & Time": "2026-03-01 10:00",
                "Event Created Date & Time": "2026-02-16 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "b@example.com",
            },
            {
                "Start Date & Time": "2026-03-05 10:00",
                "Event Created Date & Time": "2026-02-20 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "c@example.com",
            },
        ]
    )

    metrics = processor.compute_metrics(
        df=df,
        start="2026-03-01",
        end="2026-03-31",
        exclude_canceled=True,
        cost_per_training=Decimal("495"),
        evaluation_rate_percentage=Decimal("50"),
        top_topics_count=3,
        top_topics_hard_cutoff=10,
        group_topics=False,
        clean_topic_names=False,
        topic_fuzzy_match_threshold=90,
        category_reference="unused.docx",
        category_fuzzy_match_threshold=86,
    )

    assert metrics.total_sessions == 2
    assert metrics.total_evaluation_count == 1


def test_top_topics_hard_cutoff_limits_program_count(monkeypatch) -> None:
    monkeypatch.setattr(processor, "load_category_map", lambda _path: {})
    monkeypatch.setattr(
        processor,
        "categorize_topics",
        lambda topics, _category_map, _threshold: {"Uncategorized": topics},
    )

    rows = []
    for i in range(12):
        rows.append(
            {
                "Start Date & Time": f"2026-01-{i+1:02d} 10:00",
                "Event Created Date & Time": f"2026-01-{i+1:02d} 09:00",
                "Canceled": False,
                "Event Type Name": f"Topic {i+1:02d}",
                "Invitee Email": f"user{i+1:02d}@example.com",
            }
        )
    df = pd.DataFrame(rows)

    metrics = processor.compute_metrics(
        df=df,
        start="2026-01-01",
        end="2026-01-31",
        exclude_canceled=True,
        cost_per_training=Decimal("495"),
        evaluation_rate_percentage=Decimal("80"),
        top_topics_count=8,
        top_topics_hard_cutoff=10,
        group_topics=False,
        clean_topic_names=False,
        topic_fuzzy_match_threshold=90,
        category_reference="unused.docx",
        category_fuzzy_match_threshold=86,
    )

    assert len(metrics.top_topics) == 10


def test_total_sessions_counts_unique_class_dates(monkeypatch) -> None:
    monkeypatch.setattr(processor, "load_category_map", lambda _path: {})
    monkeypatch.setattr(
        processor,
        "categorize_topics",
        lambda topics, _category_map, _threshold: {"Uncategorized": topics},
    )

    df = pd.DataFrame(
        [
            {
                "Start Date & Time": "2026-02-01 10:00",
                "Event Created Date & Time": "2026-01-15 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "a@example.com",
            },
            {
                "Start Date & Time": "2026-02-01 10:00",
                "Event Created Date & Time": "2026-01-16 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "b@example.com",
            },
            {
                "Start Date & Time": "2026-02-05 10:00",
                "Event Created Date & Time": "2026-01-20 09:00",
                "Canceled": False,
                "Event Type Name": "Topic A",
                "Invitee Email": "c@example.com",
            },
        ]
    )

    metrics = processor.compute_metrics(
        df=df,
        start="2026-02-01",
        end="2026-02-28",
        exclude_canceled=True,
        cost_per_training=Decimal("495"),
        evaluation_rate_percentage=Decimal("50"),
        top_topics_count=3,
        top_topics_hard_cutoff=10,
        group_topics=False,
        clean_topic_names=False,
        topic_fuzzy_match_threshold=90,
        category_reference="unused.docx",
        category_fuzzy_match_threshold=86,
    )

    assert metrics.total_sessions == 2
