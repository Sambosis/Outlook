import webview

tasks = []  # Task list

## Compressed HTML Code to Save Space in the Article!!!
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>To-Do List</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 50px; }
        h1 { color: #333; }
        ul { list-style-type: none; padding: 0; }
        li { padding: 5px; background-color: #f9f9f9; margin-bottom: 5px; display: flex; justify-content: space-between; border-radius: 4px; border: 1px solid #ccc; }
        button { background-color: #ff4b4b; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; }
    </style>
</head>
<body>
    <h1>To-Do List</h1>
    <ul id="task-list"></ul>
    <input type="text" id="task-input" placeholder="Enter new task">
    <button onclick="addTask()">Add Task</button>

    <script>
        window.onload = () => pywebview.api.get_tasks().then(update);
        const update = tasks => document.getElementById("task-list").innerHTML = tasks.map(
            (task, i) => `<li>${task}<button onclick="remove(${i})">Remove</button></li>`
        ).join('');
        const addTask = () => pywebview.api.add_task(
            document.getElementById("task-input").value
        ).then(update);
        const remove = i => pywebview.api.remove_task(i).then(update);
    </script>
</body>
</html>
"""

class Api:
    def get_tasks(self): return tasks
    def add_task(self, task): tasks.append(task); return tasks
    def remove_task(self, index): tasks.pop(index); return tasks

if __name__ == '__main__':
    webview = webview()

    webview.config.use_qt = True
    webview.config.webview_class = 'qt'
    webview.config.debug = True
    webview.config.title = 'To-Do List App'
    webview.config.size = (400, 300)
    webview.config.resizable = True
    webview.config.frameless = False
    webview.config.fullscreen = False
    webview.config.x = 100
    webview.config.y = 100
    
    webview.create_window('To-Do List App', html=html_content, js_api=Api())
    webview.start()