# json.title

## 全体要約

json.summary

## トピック詳細

### json.topics[0].topic_title (json.topics[0].topic_time_range.start_time 〜 json.topics[0].topic_time_range.end_time<!-- optional: 存在する場合のみ表示 -->)

キーワード

- json.topics[0].topic_keywords[0]
- json.topics[0].topic_keywords[1]
- json.topics[0].topic_keywords[2]
- ...

要約

json.topics[0].topic_summary

ポイント

json.topics[0].topic_points

重要な知見

- json.topics[0].important_knowledge[0].statement
  詳細: json.topics[0].important_knowledge[0].explanation
  根拠: json.topics[0].important_knowledge[0].source
  重要度: json.topics[0].important_knowledge[0].importanceの数だけ★をつける <!-- optional -->
  タイムスタンプ: json.topics[0].important_knowledge[0].evidence.start_time 〜 json.topics[0].important_knowledge[0].evidence.end_time
- 以下続く...

専門用語

- json.topics.technical_term[0].word: json.topics.technical_term[0].explanation
  講義での定義: json.topics.technical_term[0].lecture_definition <!-- optional: 存在する場合のみ表示 -->
  誤解・混同ポイント: json.topics.technical_term[0].pitfall <!-- optional: 存在する場合のみ表示 -->
  タイムスタンプ: json.topics.technical_term[0].evidence.start_time 〜 json.topics.technical_term[0].evidence.end_time <!-- optional: 存在する場合のみ表示 -->
- 以下続く...

## 課題、提出物 <!-- optional: assignments配列が空でない場合のみセクション表示 -->

- json.assignments[0].task_title: json.assignments[0].task_detail
  提出期限: json.assignments[0].due_datetime ? json.assignments[0].due_datetime : json.assignments[0].due_date <!-- optional: いずれかが存在する場合のみ表示 -->
  タイムスタンプ: json.assignments[0].evidence.start_time 〜 json.assignments[0].evidence.end_time <!-- optional: 存在する場合のみ表示 -->
