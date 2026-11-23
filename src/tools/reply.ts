import {
    type CommsApi,
    CommsRequestError,
    getFullCommsURL,
    NOTIFY_AUDIENCES,
    type NotifyAudience,
    THREAD_ACTIONS,
    type ThreadAction,
} from '@todoist/comms-sdk'
import { z } from 'zod'
import type { CommsTool } from '../comms-tool.js'
import { type ReplyOutput, ReplyOutputSchema } from '../utils/output-schemas.js'
import { ReplyTargetTypeSchema } from '../utils/target-types.js'
import { ToolNames } from '../utils/tool-names.js'

const ArgsSchema = {
    targetType: ReplyTargetTypeSchema.describe(
        'The type of object to reply to: thread (posts a comment) or conversation (posts a message).',
    ),
    targetId: z.string().describe('The ID of the thread or conversation to reply to.'),
    content: z
        .string()
        .min(1)
        .describe(
            'The content of the reply. Markdown. Mention people with the link syntax [Name](comms-mention://USER_ID) — e.g. [Afzal](comms-mention://29367677) — never @Name or [[Name|id]], which post as literal text. Do not put "@" in the label; the client adds it. Groups use [Name](comms-group-mention://GROUP_ID), channels [#name](comms-channel://CHANNEL_ID), threads [Title](comms-thread://THREAD_ID). Resolve IDs with get-users/get-groups/list-channels first. Mentioning inline does not notify anyone. For thread replies, also pass mentioned user IDs in recipients and group IDs in groups; conversation replies notify all participants automatically and take no notification parameters.',
        ),
    recipients: z
        .array(z.number())
        .optional()
        .describe(
            'Optional array of user IDs to notify (thread replies only; ignored for conversation replies, which notify all participants automatically). If omitted with no groups and no notifyAudience, thread replies default to notifying everyone who has interacted with the thread.',
        ),
    groups: z
        .array(z.string())
        .optional()
        .describe(
            'Optional array of group IDs to notify (only for thread replies). Use get-groups to discover group IDs before passing them here.',
        ),
    notifyAudience: z
        .enum(NOTIFY_AUDIENCES)
        .optional()
        .describe(
            "Optional broader audience to notify in addition to recipients and groups (only for thread replies). 'channel' notifies everyone in the channel; 'thread' notifies everyone who has interacted with the thread.",
        ),
    threadAction: z
        .enum(THREAD_ACTIONS)
        .optional()
        .describe(
            "Thread replies only. 'close' closes the thread for everyone with this reply; it then accepts no new comments until reopened. 'reopen' reopens it. Unlike mark-done, which only archives the thread in your own inbox.",
        ),
}

/**
 * The SDK call behind each thread action and the state the reply reports.
 * Keyed by the SDK's ThreadAction, so an action the SDK adds fails to compile
 * here instead of being posted as a plain comment.
 */
const THREAD_ACTION_STEPS = {
    close: { method: 'closeThread', result: 'Closed' },
    reopen: { method: 'reopenThread', result: 'Reopened' },
} as const satisfies Record<ThreadAction, { method: keyof CommsApi['threads']; result: string }>

/**
 * A closed thread refuses new comments. The API documents no error code for
 * that, so when Comms refuses a comment (a 4xx) this reads the thread: if it
 * is closed, the error says how to post anyway. Any other case, including a
 * failed lookup, keeps the original error.
 */
async function explainRefusedComment(
    client: CommsApi,
    threadId: string,
    error: unknown,
): Promise<unknown> {
    const status = error instanceof CommsRequestError ? error.httpStatusCode : undefined
    if (status === undefined || status < 400 || status >= 500) {
        return error
    }
    const thread = await client.threads.getThread(threadId).catch(() => undefined)
    if (!thread?.closed) {
        return error
    }
    return new Error(
        `Thread ${threadId} is closed, so it does not accept new comments. To post this reply, send it again with threadAction: "reopen", which reopens the thread and posts the comment in one step.`,
    )
}

const reply = {
    name: ToolNames.REPLY,
    description:
        'Post a reply to a thread (as a comment) or conversation (as a message). Thread replies notify everyone who has interacted with the thread by default unless specific user recipients, groups, or a notifyAudience are provided. Pass threadAction to close or reopen the thread with the reply.',
    parameters: ArgsSchema,
    outputSchema: ReplyOutputSchema.shape,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    async execute(args, client) {
        const { targetType, targetId, content, recipients, groups, notifyAudience, threadAction } =
            args

        if (targetType === 'conversation' && groups !== undefined) {
            throw new Error('groups can only be used when replying to a thread.')
        }
        if (targetType === 'conversation' && notifyAudience !== undefined) {
            throw new Error('notifyAudience can only be used when replying to a thread.')
        }
        if (targetType === 'conversation' && threadAction !== undefined) {
            throw new Error(
                'threadAction can only be used when replying to a thread; conversations cannot be closed.',
            )
        }

        const groupsToNotify =
            targetType === 'thread' && groups && groups.length > 0 ? groups : undefined
        const appliedAudience: NotifyAudience | undefined =
            targetType === 'thread'
                ? (notifyAudience ??
                  (recipients === undefined && !groupsToNotify ? 'thread' : undefined))
                : undefined

        let replyId: string
        let created: Date
        let replyUrl: string

        if (targetType === 'thread') {
            const notification = {
                content,
                recipients,
                groups: groupsToNotify,
                notifyAudience: appliedAudience,
            }
            const comment = threadAction
                ? await client.threads[THREAD_ACTION_STEPS[threadAction].method]({
                      id: targetId,
                      ...notification,
                  })
                : await client.comments
                      .createComment({ threadId: targetId, ...notification })
                      .catch(async (error: unknown) => {
                          throw await explainRefusedComment(client, targetId, error)
                      })
            replyId = comment.id
            replyUrl =
                comment.url ??
                getFullCommsURL({
                    workspaceId: comment.workspaceId,
                    channelId: comment.channelId,
                    threadId: comment.threadId,
                    commentId: comment.id,
                })
            const postedValue = comment.posted
            created = postedValue
                ? typeof postedValue === 'string'
                    ? new Date(postedValue)
                    : postedValue
                : new Date()
        } else {
            const message = await client.conversationMessages.createMessage({
                conversationId: targetId,
                content,
            })
            replyId = message.id
            replyUrl =
                message.url ??
                getFullCommsURL({
                    workspaceId: message.workspaceId,
                    conversationId: message.conversationId,
                    messageId: message.id,
                })
            const postedValue = message.posted
            created = postedValue
                ? typeof postedValue === 'string'
                    ? new Date(postedValue)
                    : postedValue
                : new Date()
        }

        const lines: string[] = [
            `# Reply Posted`,
            '',
            `**Target:** ${targetType === 'thread' ? `Thread ${targetId}` : `Conversation ${targetId}`}`,
            `**Reply ID:** ${replyId}`,
            `**Created:** ${created.toISOString()}`,
            ...(threadAction
                ? [`**Thread state:** ${THREAD_ACTION_STEPS[threadAction].result}`]
                : []),
            '',
            '## Content',
            '',
            content,
        ]

        const structuredContent: ReplyStructured = {
            type: 'reply_result',
            success: true,
            targetType,
            targetId,
            replyId,
            content,
            created: created.toISOString(),
            replyUrl,
            ...(targetType === 'thread' && recipients ? { recipients } : {}),
            ...(groupsToNotify ? { groups: groupsToNotify } : {}),
            ...(appliedAudience ? { notifyAudience: appliedAudience } : {}),
            ...(threadAction ? { threadAction } : {}),
        }

        return {
            textContent: lines.join('\n'),
            structuredContent,
        }
    },
} satisfies CommsTool<typeof ArgsSchema, typeof ReplyOutputSchema.shape>

type ReplyStructured = ReplyOutput

export { reply, type ReplyStructured }
