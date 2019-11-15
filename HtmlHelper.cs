using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace MsiFileReport
{
	//----------------------------------------------------------------------------------------------------------------------------
	/// <summary>
	/// Provides HTML helper functions. (For example: getting an HTML table generated from a list of objects.)
	/// </summary>
	//----------------------------------------------------------------------------------------------------------------------------
	internal static class HtmlHelper
	{
		private static readonly string WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows).Replace('\\', '/');

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Returns an html table generated from the specified list of objects.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="list">The list of objects.</param>
		/// <param name="fxns">
		/// The properties and/or fields of each object contained in the list. (Note the order of these expressions determines 
		/// the column order in the resultant table.)
		/// </param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		public static string GetTableFromList<T>(IEnumerable<T> list, params Expression<Func<T, object>>[] fxns)
		{
			var sb = new StringBuilder();
			sb.AppendLine("<table>");
			sb.AppendLine("<tr>");
			foreach (Expression<Func<T, object>> fxn in fxns)
			{
				sb.Append("<th>");
				sb.Append(GetName(fxn));
				sb.AppendLine("</th>");
			}
			sb.AppendLine("</tr>");
			foreach (T item in list)
			{
				sb.AppendLine("<tr>");
				foreach (Expression<Func<T, object>> fxn in fxns)
				{
					sb.Append("<td>");
					object x = fxn.Compile()(item);
					if (fxn.Body.ToString() == "m.FileName")
						sb.Append($"<a href=\"file:///{WinDir}/Installer/{x}\">{x}</a>");
					else
						sb.Append(x);

					sb.AppendLine("</td>");
				}
				sb.AppendLine("</tr>");
			}
			sb.AppendLine("</table>");
			return sb.ToString();
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Gets the name of the specified expression by casting its body into either a <see cref="MemberExpression"/> or a <see 
		/// cref="UnaryExpression"/>.
		/// </summary>
		/// <typeparam name="T">The type</typeparam>
		/// <param name="expr">The expression of type T.</param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		private static string GetName<T>(Expression<Func<T, object>> expr)
		{
            if (expr.Body is MemberExpression member)
				return GetName2(member);
            return expr.Body is UnaryExpression unary ? GetName2((MemberExpression)unary.Operand) : "?+?";
		}

		//------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Gets the name from the specified member expression by reflecting on its field or property info.
		/// </summary>
		/// <param name="member">The member expression.</param>
		/// <returns></returns>
		//------------------------------------------------------------------------------------------------------------------------
		private static string GetName2(MemberExpression member)
		{
			var fieldInfo = member.Member as FieldInfo;
            if (fieldInfo != null)
                return fieldInfo.GetCustomAttribute(typeof(DescriptionAttribute)) is DescriptionAttribute fi
                    ? fi.Description
                    : fieldInfo.Name;
            var propertyInfo = member.Member as PropertyInfo;
			if (propertyInfo == null)
				return "?-?";
            return propertyInfo.GetCustomAttribute(typeof(DescriptionAttribute)) is DescriptionAttribute pi ? pi.Description : propertyInfo.Name;
		}
	}
}
