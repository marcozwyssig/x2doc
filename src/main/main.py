# main.py
from invoke import Collection, Context
import tasks

def execute_task(ns, task_name, ctx, executed=None, *args):
    if executed is None:
        executed = set()

    if task_name in executed:
        return

    try:
        task = ns[task_name]
    except KeyError:
        print(f"Task '{task_name}' not found in the collection.")
        return

    # Execute pre-tasks first
    for pre_task in task.pre:
        # pre_task is a function decorated with @task
        pre_task_name = pre_task.name
        execute_task(ns, pre_task_name, ctx, executed, *args)

    # Execute the main task
    print(f"Executing task: {task_name} with arguments {args}")
    task(ctx, *args)

    # Mark this task as executed
    executed.add(task_name)

def main():
    # Collect all tasks from tasks.py
    ns = Collection.from_module(tasks)

    # Create a context to pass to tasks
    ctx = Context()

    #task_name = "create_word_from_x2doc"
    task_name = "create_x2doc_from_word"
    
    # Example arguments to pass to the task
    #args = ['/home/zwym/Documents/eclipse/products/products/src/x2doc/product_catalog_model_tako_tm/TaKo_tm_Deployment_Analysis.x2doc', '/home/zwym/Documents/eclipse/products/products/src/x2doc/product_catalog_model_tako_tm/TaKo_tm_Deployment_Analysis.docx']
    args = ['/home/zwym/Documents/eclipse/products/products/src/x2doc/product_catalog_model_tako_tm/TaKo_tm_Deployment_Analysis.docx', '/home/zwym/Documents/eclipse/products/products/src/x2doc/product_catalog_model_tako_tm/TaKo_tm_Deployment_Analysis.x2doc']

    # Execute the specified task along with its dependencies
    execute_task(ns, task_name, ctx, None, *args)

if __name__ == '__main__':
    main()
