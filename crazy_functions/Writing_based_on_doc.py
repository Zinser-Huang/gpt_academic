from toolbox import update_ui
from toolbox import CatchException, report_exception
from toolbox import write_history_to_file, promote_file_to_downloadzone
from .crazy_utils import request_gpt_model_in_new_thread_with_ui_alive


journal = 'Nature series of journals'
writing_part = 'Discussion'
prefix_isay = f'You are expert in hydroclimatic forecasting and have been editor and professor for many years. Your task is to write {writing_part} section based on the content of a academic paper.'
require_isay = f'The language should be concise, coherent and natural and follow the academic writing style of the {journal}. The section should be divided into several paragraphs, each paragraph should not exceed 200 words. The number of paragraphs is determined by common choices in {journal}.'

def analyze_docx(file_manifest, project_folder, llm_kwargs, plugin_kwargs, chatbot, history, system_prompt, ):
    import time, os
    from docx import Document

    for index, fp in enumerate(file_manifest):
        doc = Document(fp)
        file_content = "\n".join([para.text for para in doc.paragraphs])

        i_say = f'{prefix_isay}. The academic paper is in the Word document named {os.path.relpath(fp, project_folder)}. {require_isay} The content of the document is as follows: ```{file_content}``` Please take a deep breath and start writing.'
        i_say_show_user = f'[{index}/{len(file_manifest)}] Based on the content of the Word document, please write the "{writing_part}" section in English for the file: {os.path.abspath(fp)}'
        chatbot.append((i_say_show_user, "[Local Message] waiting gpt response."))
        yield from update_ui(chatbot=chatbot, history=history) # 刷新界面

        gpt_say = yield from request_gpt_model_in_new_thread_with_ui_alive(i_say, i_say_show_user, llm_kwargs, chatbot, history=history, sys_prompt=system_prompt)
        chatbot[-1] = (i_say_show_user, gpt_say)
        history.append(i_say_show_user); history.append(gpt_say)
        yield from update_ui(chatbot=chatbot, history=history) # 刷新界面
        
    res = write_history_to_file(history)
    promote_file_to_downloadzone(res, chatbot=chatbot)
    chatbot.append(("All files have been processed?", res))
    yield from update_ui(chatbot=chatbot, history=history) # 刷新界面


@CatchException
def write_specified_part_in_docx(txt, llm_kwargs, plugin_kwargs, chatbot, history, system_prompt, user_request):
    history = []    # 清空历史,以免输入溢出
    import glob, os

    # 尝试导入依赖,如果缺少依赖,则给出安装建议
    try:
        from docx import Document
    except:
        report_exception(chatbot, history, 
                         a=f"Analyzing project: {txt}", 
                         b=f"Failed to import dependencies. Additional dependencies are required to use this module. Installation method: ```pip install python-docx```.")
        yield from update_ui(chatbot=chatbot, history=history) # 刷新界面
        return

    # 检查输入参数 
    if os.path.exists(txt):
        project_folder = txt
    else:
        if txt == "": txt = 'Empty input field'
        report_exception(chatbot, history, a = f"Analyzing project: {txt}", b = f"Unable to find local project or access denied: {txt}")
        yield from update_ui(chatbot=chatbot, history=history) # 刷新界面
        return

    file_manifest = [f for f in glob.glob(f'{project_folder}/**/*.docx', recursive=True)]
    
    if len(file_manifest) == 0:
        report_exception(chatbot, history, a = f"Analyzing project: {txt}", b = f"No .docx files found: {txt}")
        yield from update_ui(chatbot=chatbot, history=history) # 刷新界面
        return

    # 询问用户要写作的部分
    # txt = yield from get_input_async(chatbot, history, f'Please enter the name of the section you want to write, such as "Abstract", "Conclusion", etc.:')
    # plugin_kwargs['specified_part'] = 'Discussion'

    yield from analyze_docx(file_manifest, project_folder, llm_kwargs, plugin_kwargs, chatbot, history, system_prompt, )