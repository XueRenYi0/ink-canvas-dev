using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Ink_Canvas.MathGraph
{
    // =====================================================================
    // 第一部分：表达式树
    // 把公式表示成一棵"可求值"的树，比如 y = 2x+1 会变成：
    //        (+)
    //       /   \
    //    (*)     1
    //   /   \
    //  2     x
    // 求值时给 x 一个数，沿着树算上去就能得到 y
    // =====================================================================

    /// <summary>表达式树节点基类：能"给一个 x，算出一个 y"就行</summary>
    internal abstract class ExprNode
    {
        public abstract double Evaluate(double x);
    }

    /// <summary>数字节点，比如 2、3.5</summary>
    internal class NumberNode : ExprNode
    {
        private readonly double _value;
        public NumberNode(double value) { _value = value; }
        public override double Evaluate(double x) { return _value; }
    }

    /// <summary>变量节点（目前只支持 x，含参函数是远期目标）</summary>
    internal class VariableNode : ExprNode
    {
        public override double Evaluate(double x) { return x; }
    }

    /// <summary>二元运算节点：+ - * / ^</summary>
    internal class BinaryNode : ExprNode
    {
        private readonly ExprNode _left, _right;
        private readonly string _op;
        public BinaryNode(string op, ExprNode left, ExprNode right)
        {
            _op = op; _left = left; _right = right;
        }
        public string Op { get { return _op; } }
        public ExprNode Left { get { return _left; } }
        public ExprNode Right { get { return _right; } }
        public override double Evaluate(double x)
        {
            double l = _left.Evaluate(x), r = _right.Evaluate(x);
            switch (_op)
            {
                case "+": return l + r;
                case "-": return l - r;
                case "*": return l * r;
                case "/": return l / r;
                case "^": return Math.Pow(l, r);
                default: throw new FormatException("未知运算符 " + _op);
            }
        }
    }

    /// <summary>负号节点（一元运算）</summary>
    internal class NegateNode : ExprNode
    {
        private readonly ExprNode _inner;
        public NegateNode(ExprNode inner) { _inner = inner; }
        public ExprNode Inner { get { return _inner; } }
        public override double Evaluate(double x) { return -_inner.Evaluate(x); }
    }

    /// <summary>函数调用节点，比如 sin(x)、ln(x)</summary>
    internal class FunctionNode : ExprNode
    {
        private readonly string _name;
        private readonly ExprNode _arg;
        public FunctionNode(string name, ExprNode arg)
        {
            _name = name; _arg = arg;
        }
        public string Name { get { return _name; } }
        public ExprNode Argument { get { return _arg; } }
        public override double Evaluate(double x)
        {
            double v = _arg.Evaluate(x);
            switch (_name)
            {
                case "sin": return Math.Sin(v);
                case "cos": return Math.Cos(v);
                case "tan": return Math.Tan(v);
                case "cot": return 1.0 / Math.Tan(v);
                case "sec": return 1.0 / Math.Cos(v);
                case "csc": return 1.0 / Math.Sin(v);
                case "arcsin": return Math.Asin(v);
                case "arccos": return Math.Acos(v);
                case "arctan": return Math.Atan(v);
                case "ln": return Math.Log(v);
                case "log": return Math.Log10(v);
                case "lg": return Math.Log10(v);
                case "sqrt": return Math.Sqrt(v);
                case "abs": return Math.Abs(v);
                case "exp": return Math.Exp(v);
                default: throw new FormatException("未知函数 " + _name);
            }
        }
    }

    // =====================================================================
    // 第二部分：词法——把 MathML XML 转成一串记号（Token）
    // 比如 y=x+1/x 会转成：[y] [=] [x] [+] [(] [1] [/] [x] [)]
    // 之后就可以像普通表达式一样做语法分析了
    // =====================================================================

    internal enum TokenType
    {
        Number,     // 数字
        Variable,   // 变量 x
        Constant,   // 常量 π、e
        Operator,   // 运算符 + - * / ^ =
        Function,   // 函数名 sin cos ...
        LeftParen,  // (
        RightParen  // )
    }

    internal class Token
    {
        public TokenType Type;
        public string Text;     // 显示用的文本
        public double Value;    // 数字/常量的值
    }

    /// <summary>
    /// MathML → 记号流 + 表达式树 的解析器
    ///
    /// 支持的 MathML 元素（微软数学识别器会输出的那些）：
    /// mi(标识符) mn(数字) mo(运算符) mfrac(分式) msup(上标/幂)
    /// msqrt(平方根) mroot(n次根) mfenced(括号) mrow(分组)
    /// </summary>
    public static class MathMLParser
    {
        /// <summary>支持的函数名集合</summary>
        private static readonly HashSet<string> KnownFunctions = new HashSet<string>
        {
            "sin","cos","tan","cot","sec","csc",
            "arcsin","arccos","arctan",
            "ln","log","lg","sqrt","abs","exp"
        };

        /// <summary>
        /// 解析 MathML，返回"给 x 算 y"的函数（这就是画图要用的东西）
        /// 解析失败抛 FormatException，消息用中文说明原因
        /// </summary>
        /// <param name="hasAbs">输出：表达式里是否含绝对值（画图时决定要不要平滑）</param>
        public static Func<double, double> Compile(string mathml, out bool hasAbs)
        {
            ExprNode tree = ParseToTree(mathml);
            hasAbs = ContainsAbs(tree);
            return x => tree.Evaluate(x);
        }

        /// <summary>遍历表达式树，看里面有没有绝对值函数（决定画图平滑策略）</summary>
        private static bool ContainsAbs(ExprNode node)
        {
            if (node is FunctionNode fn)
                return fn.Name == "abs" || ContainsAbs(fn.Argument);
            if (node is BinaryNode b)
                return ContainsAbs(b.Left) || ContainsAbs(b.Right);
            if (node is NegateNode n)
                return ContainsAbs(n.Inner);
            return false; //数字、变量不会有 abs
        }


        /// <summary>把 MathML 转成人能看懂的表达式文本（用于界面核对）</summary>
        public static string ToPlainText(string mathml)
        {
            var tokens = TokenizeAll(mathml);
            return string.Join(" ", tokens.Select(t => t.Text));
        }

        // ---------------- 对外入口结束，下面是内部实现 ----------------

        private static ExprNode ParseToTree(string mathml)
        {
            var tokens = TokenizeAll(mathml);

            //处理等号：三种情况
            //  y=x+1 / f(x)=2x+1 → 取右侧
            //  x+1=y（罕见写法）→ 取左侧
            //  两侧都有或都没有 x → 按原来的规则兜底
            int eq = tokens.FindIndex(t => t.Type == TokenType.Operator && t.Text == "=");
            if (eq >= 0)
            {
                var left = tokens.Take(eq).ToList();
                var right = tokens.Skip(eq + 1).ToList();
                bool leftHasX = left.Any(t => t.Type == TokenType.Variable);
                bool rightHasX = right.Any(t => t.Type == TokenType.Variable);
                if (IsJustVariableX(left) && rightHasX) tokens = right;   //f(x)=...：左侧只剩孤立的 (x)
                else if (rightHasX && !leftHasX) tokens = right;          //常规 y=f(x)
                else if (leftHasX && !rightHasX) tokens = left;           //反写 x+1=y
                else if (!leftHasX && !rightHasX)
                    throw new FormatException("等式里没有变量 x，无法当作函数画图");
                //两侧都有 x（比如 x²+y²=1 的隐方程）：留给后面的解析层判断，暂会报"暂不支持"
            }

            if (!tokens.Any(t => t.Type == TokenType.Variable))
                throw new FormatException("表达式中没有变量 x");

            var parser = new TokenParser(tokens);
            ExprNode result = parser.ParseWholeExpression();
            if (parser.HasRemainingTokens())
                throw new FormatException("表达式有多余的部分：" + parser.RemainingText());
            return result;
        }

        /// <summary>
        /// 判断一段记号是否"只是孤立的变量 x"（允许外面套任意层括号）
        /// 用于识别 f(x)=、g(x)= 这种函数记号形式——此时等号左侧整体是"函数名字"，应取右侧
        /// </summary>
        private static bool IsJustVariableX(List<Token> tokens)
        {
            //剥掉所有括号，剩下的必须恰好是一个变量 x
            var stripped = tokens.Where(t => t.Type != TokenType.LeftParen && t.Type != TokenType.RightParen).ToList();
            return stripped.Count == 1 && stripped[0].Type == TokenType.Variable;
        }

        /// <summary>把 MathML 字符串解析成 XML 并遍历生成记号流</summary>
        private static List<Token> TokenizeAll(string mathml)
        {
            XDocument doc;
            try
            {
                //识别器输出的 MathML 可能带命名空间前缀 m:，XDocument 能正常处理
                doc = XDocument.Parse(mathml);
            }
            catch (Exception ex)
            {
                throw new FormatException("MathML 不是合法的 XML：" + ex.Message);
            }
            var tokens = new List<Token>();
            var root = doc.Root;
            if (root == null) throw new FormatException("MathML 内容为空");

            _insideAbs = false; //重置绝对值竖线配对状态（静态字段，防止上次解析残留）

            //根节点可能是 <math>，也可能是别的，统一从根开始递归
            TokenizeElement(root, tokens);
            return tokens;
        }

        /// <summary>递归遍历 XML 节点，把数学含义翻译成记号</summary>
        private static void TokenizeElement(XElement el, List<Token> tokens)
        {
            //忽略 XML 命名空间，只看标签本名（m:mi 和 mi 一视同仁）
            string name = el.Name.LocalName;

            switch (name)
            {
                case "math":
                case "mrow":
                case "semantics":
                case "annotation-xml":
                case "mstyle":
                    //纯容器节点：直接处理子节点
                    foreach (var child in el.Elements()) TokenizeElement(child, tokens);
                    break;

                case "mi":
                    TokenizeIdentifier(el.Value.Trim(), tokens);
                    break;

                case "mn":
                    double num;
                    if (!double.TryParse(el.Value.Trim(),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out num))
                        throw new FormatException("无法识别的数字：" + el.Value);
                    tokens.Add(new Token { Type = TokenType.Number, Text = el.Value.Trim(), Value = num });
                    break;

                case "mo":
                    TokenizeOperator(el.Value.Trim(), tokens);
                    break;

                case "mfrac":
                    //分式：(分子) / (分母)，两侧都加括号保证优先级正确
                    var parts = el.Elements().ToList();
                    if (parts.Count != 2) throw new FormatException("分式结构不完整");
                    tokens.Add(MakeParen("("));
                    foreach (var child in parts[0].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    tokens.Add(new Token { Type = TokenType.Operator, Text = "/" });
                    tokens.Add(MakeParen("("));
                    foreach (var child in parts[1].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    break;

                case "msup":
                    //上标：底^指数，例 x² → (x)^(2)
                    var supParts = el.Elements().ToList();
                    if (supParts.Count != 2) throw new FormatException("上标结构不完整");
                    tokens.Add(MakeParen("("));
                    foreach (var child in supParts[0].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    tokens.Add(new Token { Type = TokenType.Operator, Text = "^" });
                    tokens.Add(MakeParen("("));
                    foreach (var child in supParts[1].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    break;

                case "msqrt":
                    //平方根：sqrt(内容)
                    tokens.Add(new Token { Type = TokenType.Function, Text = "√" });
                    tokens.Add(MakeParen("("));
                    foreach (var child in el.Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    break;

                case "mroot":
                    //n次根：mroot(被开方数, 次数) → 被开方数^(1/次数)
                    var rootParts = el.Elements().ToList();
                    if (rootParts.Count != 2) throw new FormatException("根式结构不完整");
                    tokens.Add(MakeParen("("));
                    foreach (var child in rootParts[0].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(MakeParen(")"));
                    tokens.Add(new Token { Type = TokenType.Operator, Text = "^" });
                    tokens.Add(MakeParen("("));
                    foreach (var child in rootParts[1].Elements()) TokenizeElement(child, tokens);
                    tokens.Add(new Token { Type = TokenType.Operator, Text = "/" });
                    tokens.Add(new Token { Type = TokenType.Number, Text = "1", Value = 1 });
                    tokens.Add(MakeParen(")"));
                    break;

                case "mfenced":
                case "mo_fence":
                    //括号组：(内容)。特别地 open/close 是 | 时表示绝对值 |内容|
                    string open = (string)el.Attribute("open") ?? "(";
                    string close = (string)el.Attribute("close") ?? ")";
                    if (open == "|" && close == "|")
                    {
                        //绝对值：等价于 abs(内容)
                        tokens.Add(new Token { Type = TokenType.Function, Text = "abs" });
                        tokens.Add(MakeParen("("));
                        foreach (var child in el.Elements()) TokenizeElement(child, tokens);
                        tokens.Add(MakeParen(")"));
                    }
                    else
                    {
                        tokens.Add(MakeParen("("));
                        foreach (var child in el.Elements()) TokenizeElement(child, tokens);
                        tokens.Add(MakeParen(")"));
                    }
                    break;

                case "msub":
                case "msubsup":
                    throw new FormatException("暂不支持下标（可能包含参数或数列记号）");

                default:
                    //不认识的标签：尝试继续处理子节点（比直接报错宽容一些）
                    foreach (var child in el.Elements()) TokenizeElement(child, tokens);
                    break;
            }
        }

        /// <summary>处理标识符：x 是变量，π/e 是常量，函数记号 f/g/h 直接跳过</summary>
        private static void TokenizeIdentifier(string text, List<Token> tokens)
        {
            if (text == "x" || text == "X")
                tokens.Add(new Token { Type = TokenType.Variable, Text = "x" });
            else if (text == "π" || text == "pi" || text == "PI")
                tokens.Add(new Token { Type = TokenType.Constant, Text = "π", Value = Math.PI });
            else if (text == "e")
                tokens.Add(new Token { Type = TokenType.Constant, Text = "e", Value = Math.E });
            else if (text == "y" || text == "Y")
            {
                //y 是因变量名，画图时不需要它——真正的"="记号来自 <mo>=</mo>
                return;
            }
            else if (text == "f" || text == "g" || text == "h" ||
                     text == "F" || text == "G" || text == "H")
            {
                //函数记号 f(x)=2x+1 等价于 y=2x+1：跳过 f，后面的 (x) 解析成 x
                //（括号里恰好是单独的变量 x，ParseAtom 求值后值就是 x 本身，无需特判）
                return;
            }
            else
                throw new FormatException("暂不支持符号 \"" + text + "\"（含参数函数是后期版本的目标）");
        }

        /// <summary>处理运算符：普通符号、函数名、以及隐藏的"应用函数"符号</summary>
        private static void TokenizeOperator(string text, List<Token> tokens)
        {
            //⁡ 是 MathML 的"函数应用"隐形符号（U+2061），跳过即可
            if (text == "⁡" || text.Length == 0) return;

            //绝对值竖线：第一个 | 相当于 abs(，第二个 | 相当于 )
            //（用翻转标记配对，不支持嵌套 |a|b| 这种罕见写法）
            if (text == "|" || text == "∣")
            {
                if (!_insideAbs)
                {
                    tokens.Add(new Token { Type = TokenType.Function, Text = "abs" });
                    tokens.Add(MakeParen("("));
                }
                else
                {
                    tokens.Add(MakeParen(")"));
                }
                _insideAbs = !_insideAbs;
                return;
            }

            if (text == "+" || text == "-" || text == "*" || text == "/" || text == "^" || text == "=")
            {
                tokens.Add(new Token { Type = TokenType.Operator, Text = text });
                return;
            }

            if (KnownFunctions.Contains(text))
            {
                tokens.Add(new Token { Type = TokenType.Function, Text = text });
                return;
            }

            throw new FormatException("暂不支持符号 \"" + text + "\"");
        }

        /// <summary>绝对值竖线配对状态（每次 TokenizeAll 开始时重置）</summary>
        private static bool _insideAbs;

        private static Token MakeParen(string s)
        {
            return new Token
            {
                Type = s == "(" ? TokenType.LeftParen : TokenType.RightParen,
                Text = s
            };
        }
    }

    // =====================================================================
    // 第三部分：语法分析（递归下降法）
    // 按运算优先级从低到高逐层解析：
    // 表达式 → 加减 → 乘除(含隐式乘法 2x) → 负号 → 幂 → 原子
    // =====================================================================
    internal class TokenParser
    {
        private readonly List<Token> _tokens;
        private int _pos;

        public TokenParser(List<Token> tokens) { _tokens = tokens; }

        public bool HasRemainingTokens() { return _pos < _tokens.Count; }
        public string RemainingText()
        {
            return string.Join(" ", _tokens.Skip(_pos).Select(t => t.Text));
        }

        /// <summary>入口：解析整个表达式</summary>
        public ExprNode ParseWholeExpression() { return ParseAddSubtract(); }

        /// <summary>加减层：a + b - c ...</summary>
        private ExprNode ParseAddSubtract()
        {
            ExprNode left = ParseMultiplyDivide();
            while (true)
            {
                string op = PeekOperator("+" , "-");
                if (op == null) return left;
                _pos++; //吃掉运算符
                ExprNode right = ParseMultiplyDivide();
                left = new BinaryNode(op, left, right);
            }
        }

        /// <summary>乘除层：a * b / c，以及隐式乘法（2x、2(x+1)、x sin x）</summary>
        private ExprNode ParseMultiplyDivide()
        {
            ExprNode left = ParseUnary();
            while (true)
            {
                string op = PeekOperator("*", "/");
                if (op != null)
                {
                    _pos++;
                    ExprNode right = ParseUnary();
                    left = new BinaryNode(op, left, right);
                    continue;
                }
                //隐式乘法：下一个记号能"自己开始一个新项"却没有运算符
                if (StartsNewOperand())
                {
                    ExprNode right = ParseUnary();
                    left = new BinaryNode("*", left, right);
                    continue;
                }
                return left;
            }
        }

        /// <summary>一元负号层：-x、-(-x)</summary>
        private ExprNode ParseUnary()
        {
            string op = PeekOperator("-", "+");
            if (op == "-")
            {
                _pos++;
                return new NegateNode(ParseUnary());
            }
            if (op == "+")
            {
                _pos++; //正号没意义，直接跳过
                return ParseUnary();
            }
            return ParsePower();
        }

        /// <summary>幂层：x^2（右结合：2^3^2 = 2^(3^2)）</summary>
        private ExprNode ParsePower()
        {
            ExprNode left = ParseAtom();
            if (PeekOperator("^") != null)
            {
                _pos++;
                //指数部分允许是负号或另一个幂（右结合），所以回到 Unary 层
                ExprNode right = ParseUnary();
                return new BinaryNode("^", left, right);
            }
            return left;
        }

        /// <summary>原子层：数字、变量、常量、函数调用、括号表达式</summary>
        private ExprNode ParseAtom()
        {
            if (_pos >= _tokens.Count)
                throw new FormatException("表达式不完整（可能漏写了右边的部分）");

            Token t = _tokens[_pos];
            switch (t.Type)
            {
                case TokenType.Number:
                    _pos++;
                    return new NumberNode(t.Value);

                case TokenType.Variable:
                    _pos++;
                    return new VariableNode();

                case TokenType.Constant:
                    _pos++;
                    return new NumberNode(t.Value);

                case TokenType.Function:
                    _pos++;
                    //函数名后面跟它的参数：sin x 或 sin(x+2) 都支持
                    return new FunctionNode(t.Text == "√" ? "sqrt" : t.Text, ParseUnary());

                case TokenType.LeftParen:
                    _pos++; //吃掉 (
                    ExprNode inner = ParseAddSubtract();
                    if (_pos >= _tokens.Count || _tokens[_pos].Type != TokenType.RightParen)
                        throw new FormatException("括号不匹配（少了右括号）");
                    _pos++; //吃掉 )
                    return inner;

                default:
                    throw new FormatException("表达式此处出现了意外的符号：" + t.Text);
            }
        }

        /// <summary>看一眼当前位置是不是给定的运算符之一，是就返回该运算符</summary>
        private string PeekOperator(params string[] ops)
        {
            if (_pos >= _tokens.Count) return null;
            Token t = _tokens[_pos];
            if (t.Type != TokenType.Operator) return null;
            return ops.Contains(t.Text) ? t.Text : null;
        }

        /// <summary>当前位置是否是一个新操作数的开头（用于识别隐式乘法）</summary>
        private bool StartsNewOperand()
        {
            if (_pos >= _tokens.Count) return false;
            TokenType t = _tokens[_pos].Type;
            return t == TokenType.Number || t == TokenType.Variable ||
                   t == TokenType.Constant || t == TokenType.Function ||
                   t == TokenType.LeftParen;
        }
    }
}
