import json

def gather_questions_by_type(filenames, cognitive_type):
    results = []
    for fname in filenames:
        try:
            with open(fname, "r", encoding="utf-8") as f:
                data = json.load(f)
            for topic, subtopics in data.items():
                for subtopic, qtypes in subtopics.items():
                    for qtype, questions in qtypes.items():
                        for q in questions:
                            q_type = q.get("type", "").lower()
                            if cognitive_type.lower() in q_type:
                                results.append({
                                    "file": fname,
                                    "topic": topic,
                                    "subtopic": subtopic,
                                    "question": q.get("question", ""),
                                    "type": q_type
                                })
        except Exception as e:
            print(f"Error reading {fname}: {e}")
    return results

if __name__ == "__main__":
    files = [
        "question_templates_basic.json",
        "question_templates_intermediate.json",
        "question_templates_hard.json"
    ]
    cog_type = "inferential"  # Change to "remembering" or "applied" to test others
    found = gather_questions_by_type(files, cog_type)
    print(f"Found {len(found)} '{cog_type}' questions:")
    for q in found:
        print(f"[{q['file']}] {q['topic']} > {q['subtopic']}: {q['question']} (type: {q['type']})")