let json;

// 输入区域
const textarea = document.createElement('textarea');
textarea.classList.add('input');
document.body.appendChild(textarea);

// 操作区域
const div = document.createElement('div');
div.classList.add('wrapper');
document.body.appendChild(div);

// 解析JSON按钮
const btnJson = createBtn({
    innerHTML: '解析JSON',
    onclick: () => {
        try {
            json = JSON.parse(textarea.value);
            console.log('json is', json);
        } catch(e) {}
    }
});
div.appendChild(btnJson);

// 复制代码按钮
const btnCopy = createBtn({
    innerHTML: '复制代码',
    onclick: () => {
        textarea.select();
        document.execCommand('copy');
    }
});
div.appendChild(btnCopy);

// 清空控制台按钮
const btnClear = createBtn({
    innerHTML: '清空控制台',
    onclick: () => {
        console.clear();
    }
});
div.appendChild(btnClear);

// 清空输入区域按钮
const btnEmpty = createBtn({
    innerHTML: '清空输入区域',
    onclick: () => {
        textarea.value = '';
    }
});
div.appendChild(btnEmpty);

// 创建按钮函数
function createBtn(details) {
    const btn = document.createElement('span');
    copyProps(details, btn, ['className', 'id', 'name', 'innerText', 'innerHTML', 'title', 'onclick'])
    btn.classList.add('button');
    return btn;
}

function copyProp(obj1, obj2, prop) {obj1[prop] !== undefined && (obj2[prop] = obj1[prop]);}
function copyProps(obj1, obj2, props) {props.forEach((prop) => (copyProp(obj1, obj2, prop)));}