# This file was auto-generated by Fern from our API Definition.

from ..core.unchecked_base_model import UncheckedBaseModel
import typing
from .agent_config_override_config import AgentConfigOverrideConfig
import pydantic
from .tts_conversational_config_override_config import (
    TtsConversationalConfigOverrideConfig,
)
from ..core.pydantic_utilities import IS_PYDANTIC_V2


class ConversationConfigClientOverrideConfigInput(UncheckedBaseModel):
    agent: typing.Optional[AgentConfigOverrideConfig] = pydantic.Field(default=None)
    """
    Overrides for the agent configuration
    """

    tts: typing.Optional[TtsConversationalConfigOverrideConfig] = pydantic.Field(default=None)
    """
    Overrides for the TTS configuration
    """

    if IS_PYDANTIC_V2:
        model_config: typing.ClassVar[pydantic.ConfigDict] = pydantic.ConfigDict(extra="allow", frozen=True)  # type: ignore # Pydantic v2
    else:

        class Config:
            frozen = True
            smart_union = True
            extra = pydantic.Extra.allow
