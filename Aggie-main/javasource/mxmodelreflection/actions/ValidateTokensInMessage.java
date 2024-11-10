// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package mxmodelreflection.actions;

import mxmodelreflection.TokenReplacer;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;

/**
 * Validates if all required tokens are present in the message. The action returns a list of all the tokens that are not optional and not present in the text
 */
public class ValidateTokensInMessage extends CustomJavaAction<java.util.List<IMendixObject>>
{
	private final java.lang.String Text;
	/** @deprecated use com.mendix.utils.ListUtils.map(TokenList, com.mendix.systemwideinterfaces.core.IEntityProxy::getMendixObject) instead. */
	@java.lang.Deprecated(forRemoval = true)
	private final java.util.List<IMendixObject> __TokenList;
	private final java.util.List<mxmodelreflection.proxies.Token> TokenList;

	public ValidateTokensInMessage(
		IContext context,
		java.lang.String _text,
		java.util.List<IMendixObject> _tokenList
	)
	{
		super(context);
		this.Text = _text;
		this.__TokenList = _tokenList;
		this.TokenList = java.util.Optional.ofNullable(_tokenList)
			.orElse(java.util.Collections.emptyList())
			.stream()
			.map(tokenListElement -> mxmodelreflection.proxies.Token.initialize(getContext(), tokenListElement))
			.collect(java.util.stream.Collectors.toList());
	}

	@java.lang.Override
	public java.util.List<IMendixObject> executeAction() throws Exception
	{
		// BEGIN USER CODE
		
		return TokenReplacer.validateTokens(this.getContext(), this.Text, __TokenList);
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "ValidateTokensInMessage";
	}

	// BEGIN EXTRA CODE
	// END EXTRA CODE
}
