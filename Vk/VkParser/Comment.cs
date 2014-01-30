using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VkParser
{
    class Comment
    {
        public string UserId { get; set; }
        public string CommentLink { get; set; }
        public string CommentText { get; set; }
        public List<string> UserPhone { get; set; }
        public string UserName { get; set; }

        public Comment(string userId, string commentLink, string commentText)
        {
            this.UserId = userId;
            this.CommentLink = commentLink;
            this.CommentText = commentText;
        }
    }
}
